const assert = require('assert');
const JSZip = require('jszip');
const utils = require('../work_companion_utils.js');

function test(name, fn) {
  try {
    fn();
    console.log('✓', name);
  } catch (err) {
    console.error('✗', name);
    throw err;
  }
}

const asyncTests = [];
function testAsync(name, fn) {
  asyncTests.push({ name, fn });
}

test('DOCX-like structured blocks become stable anchored chunks', function() {
  const doc = utils.buildStructuredSource({
    id: 'doc_alpha',
    title: 'Tender.docx',
    parser: 'mammoth',
    rawContent: ''
  }, [
    { type: 'heading', text: 'Future capability', anchor: { headingLevel: 1 } },
    { type: 'paragraph', text: 'Microsoft 365 support is required.', anchor: { paragraph: 18 } },
    { type: 'table_row', text: 'Platform requirements | O365 tenant', anchor: { table: 'Platform requirements', row: 3 } }
  ]);
  assert.strictEqual(doc.chunks[0].anchor.heading, 'Future capability');
  assert.strictEqual(doc.chunks[1].anchor.paragraph, 18);
  assert.strictEqual(doc.chunks[2].anchor.table, 'Platform requirements');
  assert.strictEqual(doc.chunks[2].anchor.row, 3);
  assert.strictEqual(doc.chunks[1].id, 'doc_alpha_para18');
});

test('exact term retrieval returns matching chunk', function() {
  const doc = utils.buildStructuredSource({ id: 'doc1', title: 'One.docx' }, [
    { type: 'paragraph', text: 'The future capability includes document QA.', anchor: { paragraph: 2 } }
  ]);
  const hits = utils.rankRelevantChunks('future capability', [doc], { maxResults: 3, maxChars: 1000 });
  assert.strictEqual(hits.length, 1);
  assert.strictEqual(hits[0].matchType, 'exact');
});

test('alias retrieval matches O365 and Microsoft 365', function() {
  const doc = utils.buildStructuredSource({ id: 'doc2', title: 'Two.docx' }, [
    { type: 'paragraph', text: 'The platform integrates with Microsoft 365 services.', anchor: { paragraph: 5 } }
  ]);
  const hits = utils.rankRelevantChunks('O365 services', [doc], { maxResults: 3, maxChars: 1000 });
  assert.strictEqual(hits.length, 1);
  assert.ok(hits[0].score > 0);
});

test('table row retrieval keeps row citation anchors', function() {
  const doc = utils.buildStructuredSource({ id: 'doc3', title: 'Three.xlsx' }, [
    { type: 'table_row', text: 'Requirement | Enabled', anchor: { sheet: 'Requirements', row: 4 } }
  ]);
  const hits = utils.rankRelevantChunks('enabled requirement', [doc], { maxResults: 3, maxChars: 1000 });
  assert.strictEqual(hits[0].anchor.sheet, 'Requirements');
  assert.strictEqual(hits[0].anchor.row, 4);
  assert.strictEqual(utils.formatCitationLocation(hits[0]), 'Sheet: Requirements · Row 4');
});

test('not found retrieval does not invent citations', function() {
  const doc = utils.buildStructuredSource({ id: 'doc4', title: 'Four.docx' }, [
    { type: 'paragraph', text: 'Completely unrelated content.', anchor: { paragraph: 1 } }
  ]);
  const hits = utils.rankRelevantChunks('nonexistent tender clause', [doc], { maxResults: 3, maxChars: 1000 });
  assert.strictEqual(hits.length, 0);
});

test('multi-document retrieval retains source identity', function() {
  const docA = utils.buildStructuredSource({ id: 'docA', title: 'A.docx' }, [
    { type: 'paragraph', text: 'Azure AD is required.', anchor: { paragraph: 1 } }
  ]);
  const docB = utils.buildStructuredSource({ id: 'docB', title: 'B.docx' }, [
    { type: 'paragraph', text: 'SharePoint Online is optional.', anchor: { paragraph: 1 } }
  ]);
  const hits = utils.rankRelevantChunks('sharepoint online', [docA, docB], { maxResults: 3, maxChars: 1000 });
  assert.strictEqual(hits[0].docId, 'docB');
  assert.strictEqual(hits[0].filename, 'B.docx');
});

test('model response normalization supports wrapped shapes, mixed IDs, fallback selection, and stale protection', function() {
  const githubTopLevel = utils.normalizeModelEntries('github', utils.extractModelEntriesFromResponse('github', [
    { id: 'openai/gpt-4.1', display_name: 'GPT-4.1' },
    { modelId: 'openai/gpt-4.1-mini', friendly_name: 'GPT-4.1 Mini' },
    { name: 'openai/gpt-4.1', publisher: 'OpenAI' },
    { invalid: true }
  ]));
  assert.deepStrictEqual(githubTopLevel.map(m => m.id), ['openai/gpt-4.1', 'openai/gpt-4.1-mini']);

  const githubData = utils.extractModelEntriesFromResponse('github', { data: [{ id: 'a' }] });
  const githubModels = utils.extractModelEntriesFromResponse('github', { models: [{ id: 'b' }] });
  const githubItems = utils.extractModelEntriesFromResponse('github', { items: [{ id: 'c' }] });
  const openaiData = utils.extractModelEntriesFromResponse('openai', { data: [{ id: 'gpt-4.1' }] });
  assert.strictEqual(githubData.length, 1);
  assert.strictEqual(githubModels.length, 1);
  assert.strictEqual(githubItems.length, 1);
  assert.strictEqual(openaiData.length, 1);

  const selectState = utils.buildModelSelectState({
    provider: 'github',
    models: githubTopLevel,
    selectedModel: 'missing',
    fallbackModels: [{ id: 'openai/gpt-4.1-mini' }]
  });
  assert.strictEqual(selectState.options.length, 2);
  assert.strictEqual(selectState.options[0].value, 'openai/gpt-4.1');
  assert.strictEqual(selectState.options[1].value, 'openai/gpt-4.1-mini');
  assert.strictEqual(selectState.selectedModel, 'openai/gpt-4.1-mini');

  assert.ok(utils.shouldApplyModelRefreshResult({ requestSeq: 4, currentSeq: 4, requestProvider: 'github', activeProvider: 'github' }));
  assert.ok(!utils.shouldApplyModelRefreshResult({ requestSeq: 3, currentSeq: 4, requestProvider: 'github', activeProvider: 'github' }));
  assert.ok(!utils.shouldApplyModelRefreshResult({ requestSeq: 4, currentSeq: 4, requestProvider: 'openai', activeProvider: 'github' }));

  const identity = utils.computeModelCacheIdentity({ provider: 'github', endpoint: 'x', token: 'abc' });
  assert.ok(utils.isModelCacheFresh({ identity, cachedAt: Date.now(), models: githubTopLevel }, identity, 1000));
  assert.ok(!utils.isModelCacheFresh({ identity: 'other', cachedAt: Date.now(), models: githubTopLevel }, identity, 1000));
});

test('section export payload contains only selected section and cached template analysis', function() {
  const payload = utils.buildJobSectionExportPayload({
    id: 'job1',
    name: 'Bid response',
    notes: 'Section notes',
    messages: [
      { role: 'system', content: 'sys' },
      { role: 'user', content: 'First question' },
      { role: 'assistant', content: 'First answer' },
      { role: 'user', content: 'Ignore attached doc', isDocument: true, sourceDocId: 'doc1' }
    ]
  }, {
    sectionTitle: 'Bid response',
    exportDate: '2026-08-11',
    templateAnalysis: { templateName: 'Brand Template' }
  });
  assert.ok(payload.contentMarkdown.includes('First question'));
  assert.ok(payload.contentMarkdown.includes('First answer'));
  assert.ok(!payload.contentMarkdown.includes('Ignore attached doc'));
  assert.deepStrictEqual(payload.templateAnalysis, { templateName: 'Brand Template' });
});

testAsync('template-aware DOCX archive preserves imported package assets and profile-only fallback works', async function() {
  const baseZip = new JSZip();
  baseZip.file('[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '</Types>');
  baseZip.file('_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>');
  baseZip.file('word/document.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<w:body><w:p><w:r><w:t>Template seed body</w:t></w:r></w:p>' +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/><w:pgMar w:top="1500" w:right="1400" w:bottom="1300" w:left="1200"/></w:sectPr>' +
    '</w:body></w:document>');
  baseZip.file('word/styles.xml', '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="TemplateFont"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>');
  baseZip.file('word/theme/theme1.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="TemplateTheme"><a:themeElements><a:clrScheme name="TemplateColorScheme"/></a:themeElements></a:theme>');
  baseZip.file('word/header1.xml', '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>TEMPLATE_HEADER_MARKER</w:t></w:r></w:p></w:hdr>');
  baseZip.file('word/footer1.xml', '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>TEMPLATE_FOOTER_MARKER</w:t></w:r></w:p></w:ftr>');
  baseZip.file('word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
    '<Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
    '<Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
    '</Relationships>');

  const baseTemplateBase64 = await baseZip.generateAsync({ type: 'base64' });
  const sectionDocumentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:body><w:p><w:r><w:t>Selected Section Only</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>';

  const packaged = await utils.buildWordDocxArchive({
    JSZip,
    documentXml: sectionDocumentXml,
    stylesXml: '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    imageRelationships: [],
    imageFiles: [],
    baseTemplateBase64,
    hasTemplateProfile: true
  });

  assert.strictEqual(packaged.mode, 'template-package');
  assert.strictEqual(packaged.usedBaseTemplate, true);
  assert.strictEqual(packaged.appliedTemplateSectPr, true);

  const outZip = await JSZip.loadAsync(packaged.base64, { base64: true });
  const outDocumentXml = await outZip.file('word/document.xml').async('string');
  const outStylesXml = await outZip.file('word/styles.xml').async('string');
  const outThemeXml = await outZip.file('word/theme/theme1.xml').async('string');
  const outHeaderXml = await outZip.file('word/header1.xml').async('string');
  const outFooterXml = await outZip.file('word/footer1.xml').async('string');
  const outRelsXml = await outZip.file('word/_rels/document.xml.rels').async('string');

  assert.ok(outDocumentXml.includes('Selected Section Only'));
  assert.ok(!outDocumentXml.includes('Template seed body'));
  assert.ok(outDocumentXml.includes('w:headerReference'));
  assert.ok(outDocumentXml.includes('w:footerReference'));
  assert.ok(outStylesXml.includes('TemplateFont'));
  assert.ok(outThemeXml.includes('TemplateTheme'));
  assert.ok(outHeaderXml.includes('TEMPLATE_HEADER_MARKER'));
  assert.ok(outFooterXml.includes('TEMPLATE_FOOTER_MARKER'));
  assert.ok(outRelsXml.includes('header1.xml'));
  assert.ok(outRelsXml.includes('footer1.xml'));

  const profileOnly = await utils.buildWordDocxArchive({
    JSZip,
    documentXml: sectionDocumentXml,
    stylesXml: '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ProfileOnlyFont"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>',
    imageRelationships: [],
    imageFiles: [],
    baseTemplateBase64: '',
    hasTemplateProfile: true
  });

  assert.strictEqual(profileOnly.mode, 'profile-only');
  const profileOnlyZip = await JSZip.loadAsync(profileOnly.base64, { base64: true });
  const profileOnlyDocumentXml = await profileOnlyZip.file('word/document.xml').async('string');
  const profileOnlyStylesXml = await profileOnlyZip.file('word/styles.xml').async('string');
  assert.ok(profileOnlyDocumentXml.includes('Selected Section Only'));
  assert.ok(!profileOnlyDocumentXml.includes('Template seed body'));
  assert.ok(profileOnlyStylesXml.includes('ProfileOnlyFont'));
});

(async function runAsyncTests() {
  for (const t of asyncTests) {
    try {
      await t.fn();
      console.log('✓', t.name);
    } catch (err) {
      console.error('✗', t.name);
      throw err;
    }
  }
  console.log('All utility tests passed.');
})().catch(function(err) {
  console.error(err);
  process.exit(1);
});
