const assert = require('assert');
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

test('model normalization, selection preservation, fallback, and cache freshness work', function() {
  const models = utils.normalizeModelEntries('github', [
    { id: 'openai/gpt-4.1', friendly_name: 'GPT-4.1', publisher: 'OpenAI' },
    { id: 'openai/gpt-4.1', friendly_name: 'GPT-4.1', publisher: 'OpenAI' },
    { id: 'openai/gpt-4.1-mini', friendly_name: 'GPT-4.1 Mini', publisher: 'OpenAI' }
  ]);
  assert.strictEqual(models.length, 2);
  assert.strictEqual(utils.resolveModelSelection({ models, selectedModel: 'openai/gpt-4.1', fallbackModels: models }).reason, 'preserved');
  assert.strictEqual(utils.resolveModelSelection({ models, selectedModel: 'missing', fallbackModels: [{ id: 'openai/gpt-4.1-mini' }] }).selectedModel, 'openai/gpt-4.1-mini');
  const identity = utils.computeModelCacheIdentity({ provider: 'github', endpoint: 'x', token: 'abc' });
  assert.ok(utils.isModelCacheFresh({ identity, cachedAt: Date.now(), models }, identity, 1000));
  assert.ok(!utils.isModelCacheFresh({ identity: 'other', cachedAt: Date.now(), models }, identity, 1000));
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

console.log('All utility tests passed.');
