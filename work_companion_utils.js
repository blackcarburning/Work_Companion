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

  function extractModelEntriesFromResponse(provider, rawResponse) {
    var response = rawResponse;
    if (Array.isArray(response)) return response.slice();
    if (!response || typeof response !== "object") return [];
    if (Array.isArray(response.data)) return response.data.slice();
    if (Array.isArray(response.models)) return response.models.slice();
    if (Array.isArray(response.items)) return response.items.slice();
    if (response.result && Array.isArray(response.result.models)) return response.result.models.slice();
    if (response.result && Array.isArray(response.result.items)) return response.result.items.slice();
    if (provider === "anthropic" && response.list && Array.isArray(response.list.data)) return response.list.data.slice();
    return [];
  }

  function describeModelResponse(rawResponse) {
    if (Array.isArray(rawResponse)) {
      return { type: "array", keys: [], length: rawResponse.length };
    }
    if (!rawResponse || typeof rawResponse !== "object") {
      return { type: typeof rawResponse, keys: [], length: 0 };
    }
    var keys = Object.keys(rawResponse).slice(0, 12);
    var list = extractModelEntriesFromResponse("", rawResponse);
    return { type: "object", keys: keys, length: list.length };
  }

  function normalizeModelEntries(provider, rawModels) {
    var items = Array.isArray(rawModels) ? rawModels : [];
    var out = [];
    var seen = {};
    var idFields = ["id", "modelId", "model_id", "model", "name", "slug", "key"];
    var labelFields = ["display_name", "friendly_name", "label", "title", "name"];
    items.forEach(function(model) {
      if (!model) return;
      var id = "";
      var label = "";
      if (typeof model === "string") {
        id = model.trim();
        label = id;
      } else if (typeof model === "object") {
        for (var i = 0; i < idFields.length && !id; i++) {
          if (model[idFields[i]] != null) id = String(model[idFields[i]]).trim();
        }
        for (var j = 0; j < labelFields.length && !label; j++) {
          if (model[labelFields[j]] != null) label = String(model[labelFields[j]]).trim();
        }
      }
      if (!id) return;
      if (seen[id]) return;
      seen[id] = true;
      if (!label) label = id;
      if (model && model.publisher) label += " (" + model.publisher + ")";
      out.push({ id: id, name: label, provider: provider });
    });
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

  function buildModelSelectState(options) {
    options = options || {};
    var provider = options.provider || "";
    var models = normalizeModelEntries(provider, options.models || []);
    var resolved = resolveModelSelection({
      models: models,
      selectedModel: options.selectedModel || "",
      fallbackModels: options.fallbackModels || []
    });
    if (!resolved.selectedModel && models[0]) resolved = { selectedModel: models[0].id, reason: "first-available" };
    return {
      models: models,
      options: models.map(function(model) { return { value: model.id, label: model.name || model.id }; }),
      selectedModel: resolved.selectedModel || "",
      reason: resolved.reason || "first-available"
    };
  }

  function shouldApplyModelRefreshResult(state) {
    state = state || {};
    return state.requestSeq === state.currentSeq && state.requestProvider === state.activeProvider;
  }

  function isModelCacheFresh(record, identity, maxAgeMs) {
    if (!record || !record.identity || !record.models || !record.models.length) return false;
    if (record.identity !== identity) return false;
    if (!record.cachedAt) return false;
    return (Date.now() - record.cachedAt) <= (maxAgeMs || 15 * 60 * 1000);
  }

  var IMAGE_TYPE_TO_MIME = {
    png: "image/png",
    jpeg: "image/jpeg",
    jpg: "image/jpeg",
    gif: "image/gif",
    webp: "image/webp",
    svg: "image/svg+xml",
    bmp: "image/bmp",
    tiff: "image/tiff",
    emf: "image/x-emf",
    wmf: "image/x-wmf"
  };

  function normalizeImageType(type) {
    var t = String(type || "").toLowerCase().trim();
    if (t === "jpg") return "jpeg";
    return t;
  }

  function extensionFromPath(path) {
    var match = String(path || "").match(/\.([a-z0-9]+)(?:$|\?)/i);
    return normalizeImageType(match ? match[1] : "");
  }

  function decodeBase64ToBytes(base64Text) {
    var text = String(base64Text || "").replace(/\s+/g, "");
    if (!text) return new Uint8Array(0);
    if (typeof Buffer !== "undefined") {
      return Uint8Array.from(Buffer.from(text, "base64"));
    }
    var binary = atob(text);
    var out = new Uint8Array(binary.length);
    for (var i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
  }

  function utf8Decode(bytes) {
    if (typeof TextDecoder !== "undefined") {
      try { return new TextDecoder("utf-8").decode(bytes); } catch (err) {}
    }
    var out = "";
    for (var i = 0; i < bytes.length; i++) out += String.fromCharCode(bytes[i]);
    return out;
  }

  function detectImageTypeFromBytes(bytes) {
    var b = bytes || new Uint8Array(0);
    if (b.length >= 8 &&
        b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47 &&
        b[4] === 0x0D && b[5] === 0x0A && b[6] === 0x1A && b[7] === 0x0A) return "png";
    if (b.length >= 3 && b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return "jpeg";
    if (b.length >= 6) {
      var gifHead = String.fromCharCode(b[0], b[1], b[2], b[3], b[4], b[5]);
      if (gifHead === "GIF87a" || gifHead === "GIF89a") return "gif";
    }
    if (b.length >= 12 &&
        b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 &&
        b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return "webp";
    if (b.length >= 2 && b[0] === 0x42 && b[1] === 0x4D) return "bmp";
    if (b.length >= 4) {
      if ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2A && b[3] === 0x00) ||
          (b[0] === 0x4D && b[1] === 0x4D && b[2] === 0x00 && b[3] === 0x2A)) return "tiff";
      if (b[0] === 0xD7 && b[1] === 0xCD && b[2] === 0xC6 && b[3] === 0x9A) return "wmf";
      if (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x09 && b[3] === 0x00) return "wmf";
      if (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x00 && b[3] === 0x00 &&
          b.length >= 44 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4D && b[43] === 0x46) return "emf";
    }
    var text = utf8Decode(b.slice(0, Math.min(b.length, 512))).replace(/^\uFEFF/, "").trim();
    if (/^<\?xml[\s\S]*?<svg[\s>]/i.test(text) || /^<svg[\s>]/i.test(text)) return "svg";
    return "";
  }

  function parseDataUrl(dataUrl) {
    var match = String(dataUrl || "").match(/^data:([^;,]+)?(?:;([^,]*))?,(.*)$/i);
    if (!match) return null;
    var mimeType = String(match[1] || "").toLowerCase();
    var params = String(match[2] || "").toLowerCase();
    return {
      mimeType: mimeType,
      isBase64: params.indexOf("base64") !== -1,
      payload: match[3] || ""
    };
  }

  function imageTypeFromMime(mimeType) {
    var mime = String(mimeType || "").toLowerCase();
    if (!mime) return "";
    if (mime === "image/jpg") return "jpeg";
    if (mime === "image/svg+xml") return "svg";
    if (mime === "image/x-emf" || mime === "application/x-emf") return "emf";
    if (mime === "image/x-wmf" || mime === "application/x-msmetafile" || mime === "application/x-wmf") return "wmf";
    if (mime.indexOf("image/") === 0) return normalizeImageType(mime.slice(6));
    return "";
  }

  function getMimeForImageType(type) {
    return IMAGE_TYPE_TO_MIME[normalizeImageType(type)] || "";
  }

  function detectImageDescriptor(options) {
    options = options || {};
    var bytes = options.bytes instanceof Uint8Array ? options.bytes : new Uint8Array(0);
    if (!bytes.length && options.base64) bytes = decodeBase64ToBytes(options.base64);
    var extensionType = extensionFromPath(options.path || "");
    var signatureType = detectImageTypeFromBytes(bytes);
    var mimeType = String(options.mimeType || "").toLowerCase();
    var mimeBasedType = imageTypeFromMime(mimeType);
    var detectedType = signatureType || extensionType || mimeBasedType || "";
    return {
      path: String(options.path || ""),
      sizeBytes: bytes.length,
      mimeType: mimeType,
      extensionType: extensionType,
      signatureType: signatureType,
      mimeBasedType: mimeBasedType,
      detectedType: normalizeImageType(detectedType)
    };
  }

  function isProviderSupportedImageType(type, provider) {
    var normalized = normalizeImageType(type);
    var allowed = ["png", "jpeg", "gif", "webp"];
    if (provider && typeof provider === "object" && Array.isArray(provider.allowedTypes)) {
      allowed = provider.allowedTypes.map(normalizeImageType);
    }
    return allowed.indexOf(normalized) !== -1;
  }

  function validateProviderImageDataUrl(dataUrl, options) {
    options = options || {};
    var parsed = parseDataUrl(dataUrl);
    if (!parsed || !parsed.isBase64) {
      return { ok: false, reason: "invalid-data-url", detectedType: "", sizeBytes: 0 };
    }
    var bytes;
    try {
      bytes = decodeBase64ToBytes(parsed.payload);
    } catch (err) {
      return { ok: false, reason: "invalid-base64", detectedType: "", sizeBytes: 0 };
    }
    var descriptor = detectImageDescriptor({
      path: options.path || "",
      bytes: bytes,
      mimeType: parsed.mimeType
    });
    var mimeType = normalizeImageType(imageTypeFromMime(parsed.mimeType));
    if (!descriptor.detectedType) {
      return { ok: false, reason: "unknown-image-type", detectedType: "", sizeBytes: bytes.length };
    }
    if (mimeType && descriptor.signatureType && mimeType !== normalizeImageType(descriptor.signatureType)) {
      return { ok: false, reason: "mime-signature-mismatch", detectedType: descriptor.detectedType, sizeBytes: bytes.length };
    }
    if (!isProviderSupportedImageType(descriptor.detectedType, options.provider || options)) {
      return { ok: false, reason: "unsupported-image-type", detectedType: descriptor.detectedType, sizeBytes: bytes.length };
    }
    var canonicalMime = getMimeForImageType(descriptor.detectedType) || parsed.mimeType;
    return {
      ok: true,
      reason: "ok",
      detectedType: descriptor.detectedType,
      mimeType: canonicalMime,
      sizeBytes: bytes.length
    };
  }

  function selectProviderImageCandidate(candidates, options) {
    var list = Array.isArray(candidates) ? candidates.filter(Boolean) : [];
    if (!list.length) return { selected: null, reason: "no-candidates" };
    var primary = list[0];
    if (isProviderSupportedImageType(primary.detectedType, options)) {
      return { selected: primary, reason: "primary-supported", primary: primary };
    }
    var primaryGroup = primary.groupKey || "";
    var primaryStem = primary.basenameStem || "";
    for (var i = 1; i < list.length; i++) {
      if (!isProviderSupportedImageType(list[i].detectedType, options)) continue;
      if (primaryGroup && list[i].groupKey === primaryGroup) {
        return { selected: list[i], reason: "group-fallback", primary: primary };
      }
      if (primaryStem && list[i].basenameStem === primaryStem) {
        return { selected: list[i], reason: "basename-fallback", primary: primary };
      }
    }
    for (var j = 1; j < list.length; j++) {
      if (isProviderSupportedImageType(list[j].detectedType, options)) {
        return { selected: list[j], reason: "document-fallback", primary: primary };
      }
    }
    return { selected: primary, reason: "no-supported-fallback", primary: primary };
  }

  async function sanitizeTemplateImageForProvider(options) {
    options = options || {};
    var provider = options.provider || {};
    var source = options.source || null;
    if (!source || !source.dataUrl) {
      return { attached: false, metadataOnly: true, reason: "missing-source", source: source };
    }
    var validation = validateProviderImageDataUrl(source.dataUrl, {
      provider: provider,
      path: source.path || ""
    });
    if (validation.ok) {
      return {
        attached: true,
        metadataOnly: false,
        reason: "validated",
        source: source,
        dataUrl: source.dataUrl,
        detectedType: validation.detectedType,
        mimeType: validation.mimeType,
        sizeBytes: validation.sizeBytes
      };
    }

    var fallbackList = Array.isArray(options.fallbacks) ? options.fallbacks : [];
    for (var i = 0; i < fallbackList.length; i++) {
      var fallback = fallbackList[i];
      if (!fallback || !fallback.dataUrl) continue;
      var fallbackValidation = validateProviderImageDataUrl(fallback.dataUrl, {
        provider: provider,
        path: fallback.path || ""
      });
      if (fallbackValidation.ok) {
        return {
          attached: true,
          metadataOnly: false,
          reason: "fallback-" + validation.reason,
          source: source,
          fallback: fallback,
          dataUrl: fallback.dataUrl,
          detectedType: fallbackValidation.detectedType,
          mimeType: fallbackValidation.mimeType,
          sizeBytes: fallbackValidation.sizeBytes
        };
      }
    }

    if (typeof options.convertUnsupported === "function") {
      var converted = await options.convertUnsupported(source, validation);
      if (converted && converted.dataUrl) {
        var convertedValidation = validateProviderImageDataUrl(converted.dataUrl, {
          provider: provider,
          path: converted.path || source.path || ""
        });
        if (convertedValidation.ok) {
          return {
            attached: true,
            metadataOnly: false,
            reason: "converted-" + validation.reason,
            source: source,
            converted: converted,
            dataUrl: converted.dataUrl,
            detectedType: convertedValidation.detectedType,
            mimeType: convertedValidation.mimeType,
            sizeBytes: convertedValidation.sizeBytes
          };
        }
      }
    }

    return {
      attached: false,
      metadataOnly: true,
      reason: validation.reason,
      source: source
    };
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

  function extractLastSectPrXml(documentXml) {
    var matches = String(documentXml || "").match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g);
    if (!matches || !matches.length) return "";
    return matches[matches.length - 1];
  }

  function injectTemplateSectPr(documentXml, templateDocumentXml) {
    var generatedXml = String(documentXml || "");
    var templateSectPr = extractLastSectPrXml(templateDocumentXml);
    if (!templateSectPr) return { documentXml: generatedXml, applied: false };
    var generatedSectPrMatch = generatedXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>(?![\s\S]*<w:sectPr)/);
    if (!generatedSectPrMatch) return { documentXml: generatedXml, applied: false };
    return {
      documentXml: generatedXml.replace(generatedSectPrMatch[0], templateSectPr),
      applied: true
    };
  }

  function mergeDocumentRelationshipsXml(existingRelsXml, imageRelationships) {
    var xml = String(existingRelsXml || "");
    if (!xml) {
      xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    }
    var known = {};
    var idRegex = /Id="([^"]+)"/g;
    var idMatch;
    while ((idMatch = idRegex.exec(xml))) known[idMatch[1]] = true;
    var additions = "";
    (imageRelationships || []).forEach(function(rel) {
      if (!rel || !rel.id || !rel.target || known[rel.id]) return;
      known[rel.id] = true;
      additions += '<Relationship Id="' + rel.id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + rel.target + '"/>';
    });
    if (!additions) return xml;
    if (xml.indexOf("</Relationships>") !== -1) {
      return xml.replace("</Relationships>", additions + "</Relationships>");
    }
    return xml + additions;
  }

  function ensureContentTypesForImages(contentTypesXml, imageRelationships) {
    var xml = String(contentTypesXml || "");
    if (!xml) {
      xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '</Types>';
    }
    var extByContentType = {
      png: "image/png",
      jpeg: "image/jpeg",
      jpg: "image/jpeg",
      gif: "image/gif",
      webp: "image/webp",
      bmp: "image/bmp",
      tif: "image/tiff",
      tiff: "image/tiff",
      svg: "image/svg+xml",
      emf: "image/x-emf",
      wmf: "image/x-wmf"
    };
    var needed = {};
    (imageRelationships || []).forEach(function(rel) {
      if (!rel || !rel.target) return;
      var match = rel.target.match(/\.([a-z0-9]+)(?:$|\?)/i);
      if (!match) return;
      var ext = match[1].toLowerCase();
      if (extByContentType[ext]) needed[ext] = extByContentType[ext];
    });
    Object.keys(needed).forEach(function(ext) {
      var extensionPattern = new RegExp('Extension="' + ext.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + '"', "i");
      if (!extensionPattern.test(xml) && xml.indexOf("</Types>") !== -1) {
        xml = xml.replace("</Types>", '<Default Extension="' + ext + '" ContentType="' + needed[ext] + '"/></Types>');
      }
    });
    return xml;
  }

  async function buildWordDocxArchive(options) {
    options = options || {};
    var JSZipImpl = options.JSZip;
    if (!JSZipImpl) throw new Error("JSZip is required");

    var documentXml = String(options.documentXml || "");
    var stylesXml = String(options.stylesXml || "");
    var imageRelationships = Array.isArray(options.imageRelationships) ? options.imageRelationships.slice() : [];
    var imageFiles = Array.isArray(options.imageFiles) ? options.imageFiles.slice() : [];
    var baseTemplateBase64 = String(options.baseTemplateBase64 || "");
    var hasProfile = !!options.hasTemplateProfile;

    var mode = "default";
    var usedBaseTemplate = false;
    var appliedTemplateSectPr = false;
    var templateError = "";
    var zip = null;

    if (baseTemplateBase64) {
      try {
        zip = await JSZipImpl.loadAsync(baseTemplateBase64, { base64: true });
        usedBaseTemplate = true;
        mode = "template-package";

        var templateDocumentFile = zip.file("word/document.xml");
        if (templateDocumentFile) {
          var templateDocumentXml = await templateDocumentFile.async("string");
          var injected = injectTemplateSectPr(documentXml, templateDocumentXml);
          documentXml = injected.documentXml;
          appliedTemplateSectPr = injected.applied;
        }

        zip.file("word/document.xml", documentXml);
        if (stylesXml && !zip.file("word/styles.xml")) zip.file("word/styles.xml", stylesXml);

        var existingRelsFile = zip.file("word/_rels/document.xml.rels");
        var existingRelsXml = existingRelsFile ? await existingRelsFile.async("string") : "";
        zip.file("word/_rels/document.xml.rels", mergeDocumentRelationshipsXml(existingRelsXml, imageRelationships));

        var contentTypesFile = zip.file("[Content_Types].xml");
        var contentTypesXml = contentTypesFile ? await contentTypesFile.async("string") : "";
        zip.file("[Content_Types].xml", ensureContentTypesForImages(contentTypesXml, imageRelationships));

        imageFiles.forEach(function(imgFile) {
          if (!imgFile || !imgFile.filename || !imgFile.base64) return;
          zip.file(imgFile.filename, imgFile.base64, { base64: true });
        });
      } catch (err) {
        templateError = err && err.message ? err.message : String(err);
        zip = null;
      }
    }

    if (!zip) {
      mode = hasProfile ? "profile-only" : "default";
      var contentTypesFallbackXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '</Types>';
      contentTypesFallbackXml = ensureContentTypesForImages(contentTypesFallbackXml, imageRelationships);

      var relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>';

      var docRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>';
      docRelsXml = mergeDocumentRelationshipsXml(docRelsXml, imageRelationships);

      zip = new JSZipImpl();
      zip.file("[Content_Types].xml", contentTypesFallbackXml);
      zip.file("_rels/.rels", relsXml);
      zip.file("word/document.xml", documentXml);
      zip.file("word/styles.xml", stylesXml);
      zip.file("word/_rels/document.xml.rels", docRelsXml);
      imageFiles.forEach(function(imgFile) {
        if (!imgFile || !imgFile.filename || !imgFile.base64) return;
        zip.file(imgFile.filename, imgFile.base64, { base64: true });
      });
    }

    var base64 = await zip.generateAsync({ type: "base64" });
    return {
      base64: base64,
      mode: mode,
      usedBaseTemplate: usedBaseTemplate,
      appliedTemplateSectPr: appliedTemplateSectPr,
      templateError: templateError
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
    extractModelEntriesFromResponse: extractModelEntriesFromResponse,
    describeModelResponse: describeModelResponse,
    normalizeModelEntries: normalizeModelEntries,
    computeModelCacheIdentity: computeModelCacheIdentity,
    resolveModelSelection: resolveModelSelection,
    buildModelSelectState: buildModelSelectState,
    shouldApplyModelRefreshResult: shouldApplyModelRefreshResult,
    isModelCacheFresh: isModelCacheFresh,
    detectImageTypeFromBytes: detectImageTypeFromBytes,
    imageTypeFromMime: imageTypeFromMime,
    detectImageDescriptor: detectImageDescriptor,
    validateProviderImageDataUrl: validateProviderImageDataUrl,
    isProviderSupportedImageType: isProviderSupportedImageType,
    selectProviderImageCandidate: selectProviderImageCandidate,
    sanitizeTemplateImageForProvider: sanitizeTemplateImageForProvider,
    extensionFromPath: extensionFromPath,
    getMimeForImageType: getMimeForImageType,
    buildJobSectionExportPayload: buildJobSectionExportPayload,
    buildWordDocxArchive: buildWordDocxArchive,
    injectTemplateSectPr: injectTemplateSectPr,
    mergeDocumentRelationshipsXml: mergeDocumentRelationshipsXml,
    ensureContentTypesForImages: ensureContentTypesForImages,
    excerptText: excerptText
  };
});
