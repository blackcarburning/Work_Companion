(function(root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.WorkCompanionUtils = factory();
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function() {
  var ALIAS_GROUPS = [
    ["microsoft365", "microsoft 365", "office365", "office 365", "o365", "m365"],
    ["sharepointonline", "sharepoint online", "spo"],
    ["azuread", "azure ad", "entra", "entra id", "microsoft entra"]
  ];

  function simpleHash(input) {
    var str = String(input || "");
    var hash = 5381;
    for (var i = 0; i < str.length; i++) {
      hash = ((hash << 5) + hash) + str.charCodeAt(i);
      hash = hash >>> 0;
    }
    return hash.toString(16);
  }

  function sanitizeFilename(name) {
    return String(name || "export")
      .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_+|_+$/g, "") || "export";
  }

  function normalizeWhitespace(text) {
    return String(text || "").replace(/\r\n?/g, "\n").replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim();
  }

  function applyAliasNormalization(text) {
    var normalized = " " + String(text || "").toLowerCase() + " ";
    ALIAS_GROUPS.forEach(function(group) {
      var canonical = group[0];
      for (var i = 1; i < group.length; i++) {
        var alias = group[i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s+");
        normalized = normalized.replace(new RegExp("(^|[^a-z0-9])" + alias + "([^a-z0-9]|$)", "gi"), "$1" + canonical + " $2");
      }
    });
    return normalized;
  }

  function normalizeRetrievalText(text) {
    var normalized = applyAliasNormalization(text || "");
    normalized = normalized
      .replace(/[’']/g, "")
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    return normalized;
  }

  function expandAliasToken(token) {
    var out = [token];
    ALIAS_GROUPS.forEach(function(group) {
      if (group.indexOf(token) !== -1) {
        group.forEach(function(alias) {
          var aliasToken = normalizeRetrievalText(alias);
          if (aliasToken && out.indexOf(aliasToken) === -1) out.push(aliasToken);
        });
      }
    });
    return out;
  }

  function tokenizeForRetrieval(text) {
    var normalized = normalizeRetrievalText(text);
    if (!normalized) return [];
    var seen = {};
    var out = [];
    normalized.split(" ").forEach(function(token) {
      if (!token) return;
      expandAliasToken(token).forEach(function(expanded) {
        if (!expanded || seen[expanded]) return;
        seen[expanded] = true;
        out.push(expanded);
      });
    });
    return out;
  }

  function makeChunkId(docId, block, partIndex) {
    var anchorBits = [
      block.anchor && block.anchor.page ? "p" + block.anchor.page : "",
      block.anchor && block.anchor.sheet ? "s" + block.anchor.sheet : "",
      block.anchor && block.anchor.paragraph != null ? "para" + block.anchor.paragraph : "",
      block.anchor && block.anchor.listItem != null ? "li" + block.anchor.listItem : "",
      block.anchor && block.anchor.table ? "tbl" + sanitizeFilename(block.anchor.table).toLowerCase() : "",
      block.anchor && block.anchor.row != null ? "row" + block.anchor.row : "",
      partIndex != null ? "part" + partIndex : ""
    ].filter(Boolean).join("_");
    return sanitizeFilename(docId + "_" + (anchorBits || ("chunk_" + (block.index || 0)))).toLowerCase();
  }

  function excerptText(text, maxChars) {
    var clean = normalizeWhitespace(text || "");
    var limit = maxChars || 220;
    if (clean.length <= limit) return clean;
    return clean.slice(0, limit - 1) + "…";
  }

  function cloneAnchor(anchor, heading) {
    var next = {};
    Object.keys(anchor || {}).forEach(function(key) {
      if (anchor[key] != null && anchor[key] !== "") next[key] = anchor[key];
    });
    if (heading && !next.heading) next.heading = heading;
    return next;
  }

  function normalizeBlock(block, index, lastHeading) {
    var text = normalizeWhitespace(block && block.text || "");
    var type = block && block.type || "paragraph";
    var heading = block && block.heading || lastHeading || "";
    var anchor = cloneAnchor(block && block.anchor || {}, heading);
    if (type === "heading") {
      heading = text || heading;
      anchor.heading = heading;
    }
    return {
      index: index,
      type: type,
      text: text,
      heading: heading,
      anchor: anchor
    };
  }

  function buildStructuredSource(meta, blocks, options) {
    options = options || {};
    meta = meta || {};
    var maxChunkChars = options.maxChunkChars || 900;
    var overlapChars = options.overlapChars || 140;
    var normalizedBlocks = [];
    var lastHeading = "";
    (blocks || []).forEach(function(block, index) {
      var next = normalizeBlock(block, index, lastHeading);
      if (!next.text) return;
      if (next.type === "heading") lastHeading = next.text;
      else if (next.heading) lastHeading = next.heading;
      normalizedBlocks.push(next);
    });

    var stats = {
      headingCount: 0,
      paragraphCount: 0,
      listCount: 0,
      tableCount: 0,
      tableRowCount: 0,
      pageCount: 0,
      sheetCount: 0,
      chunkCount: 0
    };

    var seenPages = {};
    var seenSheets = {};
    var seenTables = {};
    normalizedBlocks.forEach(function(block) {
      if (block.type === "heading") stats.headingCount++;
      else if (block.type === "paragraph") stats.paragraphCount++;
      else if (block.type === "list_item") stats.listCount++;
      else if (block.type === "table_row") stats.tableRowCount++;
      if (block.anchor && block.anchor.table) seenTables[block.anchor.table] = true;
      if (block.anchor && block.anchor.page != null) seenPages[block.anchor.page] = true;
      if (block.anchor && block.anchor.sheet) seenSheets[block.anchor.sheet] = true;
    });
    stats.tableCount = Object.keys(seenTables).length;
    stats.pageCount = Object.keys(seenPages).length;
    stats.sheetCount = Object.keys(seenSheets).length;

    var chunks = [];
    normalizedBlocks.forEach(function(block) {
      if (block.text.length <= maxChunkChars) {
        chunks.push({
          id: makeChunkId(meta.id || "doc", block),
          docId: meta.id || "doc",
          filename: meta.title || "",
          heading: block.heading || "",
          text: block.text,
          normalizedText: normalizeRetrievalText(block.text),
          excerpt: excerptText(block.text),
          anchor: cloneAnchor(block.anchor, block.heading),
          matchType: "best-match",
          parser: meta.parser || "",
          parserVersion: meta.parserVersion || ""
        });
        return;
      }

      var start = 0;
      var partIndex = 0;
      while (start < block.text.length) {
        var end = Math.min(block.text.length, start + maxChunkChars);
        var slice = block.text.slice(start, end).trim();
        if (!slice) break;
        var anchor = cloneAnchor(block.anchor, block.heading);
        anchor.part = partIndex + 1;
        chunks.push({
          id: makeChunkId(meta.id || "doc", block, partIndex + 1),
          docId: meta.id || "doc",
          filename: meta.title || "",
          heading: block.heading || "",
          text: slice,
          normalizedText: normalizeRetrievalText(slice),
          excerpt: excerptText(slice),
          anchor: anchor,
          matchType: "best-match",
          parser: meta.parser || "",
          parserVersion: meta.parserVersion || ""
        });
        if (end >= block.text.length) break;
        start = Math.max(end - overlapChars, start + 1);
        partIndex++;
      }
    });
    stats.chunkCount = chunks.length;

    return {
      id: meta.id || "",
      title: meta.title || "",
      type: meta.type || "uploaded_file",
      parser: meta.parser || "",
      parserVersion: meta.parserVersion || "",
      rawContent: meta.rawContent || normalizedBlocks.map(function(block) { return block.text; }).join("\n\n"),
      normalizedContent: normalizeWhitespace(meta.normalizedContent || meta.rawContent || normalizedBlocks.map(function(block) { return block.text; }).join("\n\n")),
      blocks: normalizedBlocks,
      chunks: chunks,
      stats: stats,
      createdAt: meta.createdAt || new Date().toISOString()
    };
  }

  function fallbackTextToBlocks(text) {
    var paragraph = 0;
    return normalizeWhitespace(text || "").split(/\n{2,}/).map(function(part) {
      var trimmed = part.trim();
      if (!trimmed) return null;
      paragraph++;
      return {
        type: "paragraph",
        text: trimmed,
        anchor: { paragraph: paragraph }
      };
    }).filter(Boolean);
  }

  function ensureStructuredSource(doc, options) {
    if (doc && Array.isArray(doc.chunks) && doc.chunks.length) return doc;
    var blocks = doc && Array.isArray(doc.blocks) && doc.blocks.length ? doc.blocks : fallbackTextToBlocks(doc && (doc.normalizedContent || doc.rawContent || ""));
    return buildStructuredSource({
      id: doc && doc.id || "",
      title: doc && doc.title || "",
      type: doc && doc.type || "uploaded_file",
      parser: doc && doc.parser || "legacy",
      parserVersion: doc && doc.parserVersion || "",
      rawContent: doc && doc.rawContent || "",
      normalizedContent: doc && doc.normalizedContent || "",
      createdAt: doc && doc.createdAt || ""
    }, blocks, options);
  }

  function scoreChunk(queryText, chunk) {
    var normalizedQuery = normalizeRetrievalText(queryText);
    var queryTokens = tokenizeForRetrieval(queryText);
    if (!queryTokens.length) return { score: 0, overlap: 0, exact: false };
    var chunkTokens = tokenizeForRetrieval(chunk.text || "");
    var tokenSet = {};
    chunkTokens.forEach(function(token) { tokenSet[token] = (tokenSet[token] || 0) + 1; });
    var overlap = 0;
    queryTokens.forEach(function(token) {
      if (tokenSet[token]) overlap++;
    });
    var exact = normalizedQuery && (chunk.normalizedText || normalizeRetrievalText(chunk.text)).indexOf(normalizedQuery) !== -1;
    var density = overlap ? overlap / Math.max(chunkTokens.length, 1) : 0;
    var score = (exact ? 25 : 0) + (overlap * 8) + Math.round(density * 100) / 10;
    return { score: score, overlap: overlap, exact: exact };
  }

  function rankRelevantChunks(question, documents, options) {
    options = options || {};
    var maxResults = options.maxResults || 6;
    var maxPerDoc = options.maxPerDoc || 3;
    var maxChars = options.maxChars || 5000;
    var ranked = [];

    (documents || []).forEach(function(doc) {
      var structured = ensureStructuredSource(doc, options);
      var perDoc = [];
      (structured.chunks || []).forEach(function(chunk) {
        var scoreInfo = scoreChunk(question, chunk);
        if (!scoreInfo.score) return;
        perDoc.push({
          docId: structured.id,
          filename: structured.title,
          chunkId: chunk.id,
          text: chunk.text,
          excerpt: chunk.excerpt || excerptText(chunk.text),
          heading: chunk.heading || "",
          anchor: cloneAnchor(chunk.anchor, chunk.heading),
          score: scoreInfo.score,
          overlap: scoreInfo.overlap,
          matchType: scoreInfo.exact ? "exact" : "best-match"
        });
      });
      perDoc.sort(function(a, b) { return b.score - a.score || a.chunkId.localeCompare(b.chunkId); });
      ranked = ranked.concat(perDoc.slice(0, maxPerDoc));
    });

    ranked.sort(function(a, b) {
      return b.score - a.score || a.docId.localeCompare(b.docId) || a.chunkId.localeCompare(b.chunkId);
    });

    var selected = [];
    var charBudget = 0;
    for (var i = 0; i < ranked.length && selected.length < maxResults; i++) {
      var item = ranked[i];
      if (charBudget + item.text.length > maxChars) continue;
      selected.push(item);
      charBudget += item.text.length;
    }
    return selected;
  }

  function formatCitationLocation(citation) {
    citation = citation || {};
    var anchor = citation.anchor || {};
    var parts = [];
    if (anchor.heading) parts.push("Section: " + anchor.heading);
    if (anchor.page != null) parts.push("Page " + anchor.page);
    if (anchor.sheet) parts.push("Sheet: " + anchor.sheet + (anchor.row != null ? " · Row " + anchor.row : ""));
    else if (anchor.table) parts.push("Table: " + anchor.table + (anchor.row != null ? " · Row " + anchor.row : ""));
    else if (anchor.listItem != null) parts.push("List item " + anchor.listItem);
    else if (anchor.paragraph != null) parts.push("Paragraph " + anchor.paragraph);
    if (anchor.part != null) parts.push("Part " + anchor.part);
    return parts.join(" · ");
  }

  function buildEvidenceContext(question, results) {
    var lines = [
      "## Retrieved source evidence",
      "Question: " + normalizeWhitespace(question || ""),
      "Use only the evidence below for citations. If evidence is absent, say so and return citations: [].",
      "Never claim DOCX page numbers unless a page anchor is explicitly provided."
    ];
    (results || []).forEach(function(item, index) {
      lines.push("");
      lines.push("### Evidence " + (index + 1));
      lines.push("source_id: " + item.docId);
      lines.push("chunk_id: " + item.chunkId);
      lines.push("filename: " + item.filename);
      if (item.heading) lines.push("heading: " + item.heading);
      if (item.anchor) lines.push("anchor: " + formatCitationLocation(item));
      lines.push("match_type: " + item.matchType);
      lines.push("relevance_score: " + item.score.toFixed(2));
      lines.push("excerpt:");
      lines.push(item.text);
    });
    lines.push("");
    lines.push("Respond as JSON only with keys answer, confidence, citations[].");
    return lines.join("\n");
  }

  function parseStructuredAnswer(rawText) {
    var text = String(rawText || "").trim();
    if (!text) return null;
    var cleaned = text
      .replace(/^```(?:json)?\s*/i, "")
      .replace(/\s*```$/, "")
      .trim();
    if (cleaned.charAt(0) !== "{") return null;
    try {
      var parsed = JSON.parse(cleaned);
      if (!parsed || typeof parsed.answer !== "string") return null;
      parsed.confidence = typeof parsed.confidence === "string" ? parsed.confidence : "unknown";
      parsed.citations = Array.isArray(parsed.citations) ? parsed.citations.map(function(item) {
        var citation = item || {};
        return {
          sourceId: citation.sourceId || citation.source_id || "",
          chunkId: citation.chunkId || citation.chunk_id || "",
          filename: citation.filename || "",
          heading: citation.heading || citation.section || "",
          anchor: citation.anchor || {
            heading: citation.heading || citation.section || "",
            paragraph: citation.paragraph,
            listItem: citation.listItem || citation.list_item,
            table: citation.table,
            row: citation.row,
            page: citation.page,
            sheet: citation.sheet
          },
          excerpt: normalizeWhitespace(citation.excerpt || ""),
          matchType: citation.matchType || citation.match_type || "best-match"
        };
      }).filter(function(item) { return item.sourceId || item.chunkId || item.excerpt; }) : [];
      return parsed;
    } catch (err) {
      return null;
    }
  }

  function normalizeModelEntries(provider, rawModels) {
    var items = Array.isArray(rawModels) ? rawModels : [];
    var out = [];
    var seen = {};
    items.forEach(function(model) {
      if (!model) return;
      var id = String(model.id || model.name || model.model || "").trim();
      if (!id || seen[id]) return;
      seen[id] = true;
      var label = String(model.name || model.display_name || model.friendly_name || id).trim();
      if (model.publisher) label += " (" + model.publisher + ")";
      out.push({ id: id, name: label, provider: provider });
    });
    out.sort(function(a, b) { return a.name.localeCompare(b.name); });
    return out;
  }

  function computeModelCacheIdentity(config) {
    config = config || {};
    return simpleHash([
      config.provider || "",
      config.endpoint || "",
      config.projectId || "",
      config.token || ""
    ].join("|"));
  }

  function resolveModelSelection(options) {
    options = options || {};
    var models = options.models || [];
    var fallbackModels = options.fallbackModels || [];
    var modelIds = models.map(function(model) { return model.id; });
    var selected = options.selectedModel || "";
    if (selected && modelIds.indexOf(selected) !== -1) {
      return { selectedModel: selected, reason: "preserved" };
    }
    for (var i = 0; i < fallbackModels.length; i++) {
      if (modelIds.indexOf(fallbackModels[i].id) !== -1) {
        return { selectedModel: fallbackModels[i].id, reason: "fallback" };
      }
    }
    return { selectedModel: models[0] ? models[0].id : "", reason: "first-available" };
  }

  function isModelCacheFresh(record, identity, maxAgeMs) {
    if (!record || !record.identity || !record.models || !record.models.length) return false;
    if (record.identity !== identity) return false;
    if (!record.cachedAt) return false;
    return (Date.now() - record.cachedAt) <= (maxAgeMs || 15 * 60 * 1000);
  }

  function buildJobSectionExportPayload(job, options) {
    options = options || {};
    job = job || {};
    var messages = Array.isArray(job.messages) ? job.messages.slice(1) : [];
    var transcript = [];
    messages.forEach(function(message) {
      if (!message || message.isDocument || message.isManualContent || message.isImage) return;
      transcript.push({
        role: message.role,
        content: String(message.content || "").trim(),
        excluded: !!message.excluded
      });
    });
    var exportDate = options.exportDate || new Date().toISOString().slice(0, 10);
    var sectionTitle = options.sectionTitle || job.name || "Job Section";
    var parts = ["# " + sectionTitle];
    parts.push("");
    parts.push("- Job title: " + (job.name || ""));
    parts.push("- Section title: " + sectionTitle);
    parts.push("- Export date: " + exportDate);
    if (job.notes) {
      parts.push("");
      parts.push("## Notes");
      parts.push("");
      parts.push(String(job.notes).trim());
    }
    if (transcript.length) {
      parts.push("");
      parts.push("## Conversation");
      transcript.forEach(function(entry) {
        parts.push("");
        parts.push("### " + (entry.role === "assistant" ? "Assistant" : "You"));
        parts.push(entry.content || "[No content]");
      });
    }
    return {
      jobId: job.id || "",
      jobName: job.name || "",
      sectionTitle: sectionTitle,
      exportDate: exportDate,
      notes: String(job.notes || ""),
      transcript: transcript,
      contentMarkdown: parts.join("\n"),
      templateAnalysis: options.templateAnalysis || null
    };
  }

  return {
    ALIAS_GROUPS: ALIAS_GROUPS,
    sanitizeFilename: sanitizeFilename,
    normalizeWhitespace: normalizeWhitespace,
    normalizeRetrievalText: normalizeRetrievalText,
    tokenizeForRetrieval: tokenizeForRetrieval,
    buildStructuredSource: buildStructuredSource,
    ensureStructuredSource: ensureStructuredSource,
    fallbackTextToBlocks: fallbackTextToBlocks,
    rankRelevantChunks: rankRelevantChunks,
    formatCitationLocation: formatCitationLocation,
    buildEvidenceContext: buildEvidenceContext,
    parseStructuredAnswer: parseStructuredAnswer,
    normalizeModelEntries: normalizeModelEntries,
    computeModelCacheIdentity: computeModelCacheIdentity,
    resolveModelSelection: resolveModelSelection,
    isModelCacheFresh: isModelCacheFresh,
    buildJobSectionExportPayload: buildJobSectionExportPayload,
    excerptText: excerptText
  };
});
