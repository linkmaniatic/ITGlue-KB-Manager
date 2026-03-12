var ITGlueImportDocuments = Class.create();
ITGlueImportDocuments.prototype = {
  initialize: function (mode) {
    this.debug = (mode === "debug");
  },

  _log: function (level, message) {
    if (!this.debug) return;
    var entry = "ITGlue-KBImport: " + message;
    switch (level) {
      case "warn":
        gs.warn(entry);
        break;
      case "error":
        gs.error(entry);
        break;
      default:
        gs.info(entry);
        break;
    }
  },

  _logSeparator: function (label) {
    this._log("info", "----------------" + label + "----------------");
  },

  importDocuments: function (request) {
    this._logSeparator("Initialize importDocuments");

    var body = request.body.data;

    if (!body || !Array.isArray(body.document) || body.document.length === 0) {
      this._log(
        "error",
        "Invalid payload received — document array is missing or empty.",
      );
      return {
        status: "error",
        message: "Invalid payload: document array is missing or empty.",
      };
    }

    var documents = body.document;
    this._log("info", "Payload received — document count: " + documents.length);

    var results = [];

    for (var i = 0; i < documents.length; i++) {
      var doc = documents[i];
      var companyName = String(doc.companyname || "").trim();
      var documentName = String(doc.documentname || "").trim();
      var docContent = String(doc.documentcontent || "");
      var attachments = Array.isArray(doc.attachments) ? doc.attachments : [];

      this._log(
        "info",
        "[" +
          (i + 1) +
          "/" +
          documents.length +
          '] Processing document: "' +
          documentName +
          '" | Company: "' +
          companyName +
          '" | Attachments: ' +
          attachments.length,
      );

      var docResult = {
        companyname: companyName,
        documentname: documentName,
        status: "success",
        kb_sys_id: "",
        article_sys_id: "",
        attachments: [],
      };

      try {
        var kbResult = this._getOrCreateKnowledgeBase(companyName);
        var kbSysId = kbResult.kbSysId;
        var kbDomain = kbResult.domain;
        docResult.kb_sys_id = kbSysId;

        var gsa = new GlideSysAttachment();
        var grKb = new GlideRecord("kb_knowledge_base");
        grKb.get(kbSysId);

        this._log(
          "info",
          '  Creating Knowledge Article: "' + documentName + '"',
        );
        var grArticle = new GlideRecord("kb_knowledge");
        grArticle.initialize();
        grArticle.setValue("short_description", documentName);
        grArticle.setValue("kb_knowledge_base", kbSysId);
        grArticle.setValue("workflow_state", "draft");
        grArticle.setValue("text", docContent);
        var articleSysId = grArticle.insert();
        docResult.article_sys_id = articleSysId;
        var currentArticleObj = new GlideRecord("kb_knowledge");
        currentArticleObj.get(articleSysId);
        if (kbDomain) {
          currentArticleObj.setValue("sys_domain", kbDomain);
          currentArticleObj.update();
        }
        this._log(
          "info",
          "  Knowledge Article created — sys_id: " + articleSysId,
        );

        for (var j = 0; j < attachments.length; j++) {
          var att = attachments[j];
          this._log(
            "info",
            "  Creating attachment [" +
              (j + 1) +
              "/" +
              attachments.length +
              ']: "' +
              att.filename +
              '" (' +
              att.contenttype +
              ")",
          );

          var attSysId = gsa.writeBase64(
            currentArticleObj,
            att.filename,
            att.contenttype,
            att.content,
          );

          if (attSysId) {
            this._log("info", "  Attachment created — sys_id: " + attSysId);
          } else {
            this._log(
              "warn",
              '  Attachment "' +
                att.filename +
                '" was written but returned no sys_id.',
            );
          }

          docResult.attachments.push({
            filename: att.filename,
            sys_id: attSysId,
          });
        }

        currentArticleObj.setValue("workflow_state", "published");
        currentArticleObj.update();

        this._log(
          "info",
          '  Document processed successfully: "' + documentName + '"',
        );
      } catch (e) {
        docResult.status = "error";
        docResult.message = e.message || String(e);
        this._log(
          "error",
          'Failed processing document "' +
            documentName +
            '" for company "' +
            companyName +
            '": ' +
            docResult.message,
        );
      }

      this._log(
        "info",
        'Finished processing document: "' +
          documentName +
          '" | Status: ' +
          docResult.status.toUpperCase() +
          " | Article sys_id: " +
          docResult.article_sys_id +
          " | Attachments: " +
          docResult.attachments.length,
      );

      results.push(docResult);
    }

    this._log(
      "info",
      "Batch complete — total: " +
        documents.length +
        " | success: " +
        results.filter(function (r) {
          return r.status === "success";
        }).length +
        " | error: " +
        results.filter(function (r) {
          return r.status === "error";
        }).length,
    );
    this._logSeparator("Finalized importDocuments");
  },

  importLargeDocument: function (request) {
    this._logSeparator("Initialize importLargeDocument");

    var stream = request.body.dataStream;
    var content = "";
    try {
      var reader = new GlideTextReader(stream);
      var line;
      while ((line = reader.readLine()) !== null) {
        content += line + "\n";
      }
    } catch (e) {
      gs.error("ITGlue-KBImport: Stream error: " + e);
      return { status: "error", message: "Failed to read stream: " + String(e) };
    } finally {
      stream.close();
    }

    var body;
    try {
      body = JSON.parse(content);
    } catch (e) {
      this._log("error", "Failed to parse JSON from stream: " + e);
      return { status: "error", message: "Invalid JSON in stream: " + String(e) };
    }

    if (!body || !Array.isArray(body.document) || body.document.length === 0) {
      this._log("error", "Invalid payload — document array is missing or empty.");
      return { status: "error", message: "Invalid payload: document array is missing or empty." };
    }

    this._log("info", "Large document stream parsed — document count: " + body.document.length);

    var documents = body.document;
    var results = [];

    for (var i = 0; i < documents.length; i++) {
      var doc = documents[i];
      var companyName = String(doc.companyname || "").trim();
      var documentName = String(doc.documentname || "").trim();
      var docContent = String(doc.documentcontent || "");
      var attachments = Array.isArray(doc.attachments) ? doc.attachments : [];

      this._log(
        "info",
        "[" + (i + 1) + "/" + documents.length + '] Processing large document: "' + documentName + '" | Company: "' + companyName + '" | Attachments: ' + attachments.length,
      );

      var docResult = {
        companyname: companyName,
        documentname: documentName,
        status: "success",
        kb_sys_id: "",
        article_sys_id: "",
        attachments: [],
      };

      try {
        var kbResult = this._getOrCreateKnowledgeBase(companyName);
        var kbSysId = kbResult.kbSysId;
        var kbDomain = kbResult.domain;
        docResult.kb_sys_id = kbSysId;

        var gsa = new GlideSysAttachment();

        this._log("info", '  Creating Knowledge Article: "' + documentName + '"');
        var grArticle = new GlideRecord("kb_knowledge");
        grArticle.initialize();
        grArticle.setValue("short_description", documentName);
        grArticle.setValue("kb_knowledge_base", kbSysId);
        grArticle.setValue("workflow_state", "draft");
        grArticle.setValue("text", docContent);
        var articleSysId = grArticle.insert();
        docResult.article_sys_id = articleSysId;
        var currentArticleObj = new GlideRecord("kb_knowledge");
        currentArticleObj.get(articleSysId);
        if (kbDomain) {
          currentArticleObj.setValue("sys_domain", kbDomain);
          currentArticleObj.update();
        }
        this._log("info", "  Knowledge Article created — sys_id: " + articleSysId);

        for (var j = 0; j < attachments.length; j++) {
          var att = attachments[j];
          this._log(
            "info",
            "  Creating attachment [" + (j + 1) + "/" + attachments.length + ']: "' + att.filename + '" (' + att.contenttype + ")",
          );
          var attSysId = gsa.writeBase64(currentArticleObj, att.filename, att.contenttype, att.content);
          if (attSysId) {
            this._log("info", "  Attachment created — sys_id: " + attSysId);
          } else {
            this._log("warn", '  Attachment "' + att.filename + '" was written but returned no sys_id.');
          }
          docResult.attachments.push({ filename: att.filename, sys_id: attSysId });
        }

        currentArticleObj.setValue("workflow_state", "published");
        currentArticleObj.update();
        this._log("info", '  Large document processed successfully: "' + documentName + '"');
      } catch (e) {
        docResult.status = "error";
        docResult.message = e.message || String(e);
        this._log("error", 'Failed processing large document "' + documentName + '" for company "' + companyName + '": ' + docResult.message);
      }

      this._log(
        "info",
        'Finished processing large document: "' + documentName + '" | Status: ' + docResult.status.toUpperCase() + " | Article sys_id: " + docResult.article_sys_id + " | Attachments: " + docResult.attachments.length,
      );

      results.push(docResult);
    }

    this._log(
      "info",
      "Large document batch complete — total: " + documents.length +
        " | success: " + results.filter(function (r) { return r.status === "success"; }).length +
        " | error: " + results.filter(function (r) { return r.status === "error"; }).length,
    );
    this._logSeparator("Finalized importLargeDocument");
    return results;
  },

  _getOrCreateKnowledgeBase: function (companyName) {
    var gr = new GlideRecord("kb_knowledge_base");
    gr.addQuery("title", companyName);
    gr.setLimit(1);
    gr.query();

    if (gr.next()) {
      this._log(
        "info",
        '  Knowledge Base found for "' +
          companyName +
          '" — sys_id: ' +
          gr.getUniqueValue(),
      );
      return {
        kbSysId: gr.getUniqueValue(),
        domain: gr.getValue("sys_domain"),
      };
    }

    this._log(
      "info",
      '  Knowledge Base not found for "' +
        companyName +
        '" — creating new one.',
    );
    var domain = this.checkCompany(null, companyName);
    var grNew = new GlideRecord("kb_knowledge_base");
    grNew.initialize();
    grNew.setValue("title", companyName);
    grNew.setValue("description", "Imported from ITGlue for " + companyName);
    var newKbSysId = grNew.insert();
    var newKbObj = new GlideRecord("kb_knowledge_base");
    newKbObj.get(newKbSysId);
    if (domain) {
      newKbObj.setValue("sys_domain", domain);
      newKbObj.update();
      this._log(
        "info",
        '  Domain resolved for "' + companyName + '" — sys_domain: ' + domain,
      );
    } else {
      this._log(
        "warn",
        '  Company "' +
          companyName +
          '" not found in core_company — KB created without domain assignment.',
      );
    }
    this._log(
      "info",
      '  Knowledge Base created for "' +
        companyName +
        '" — sys_id: ' +
        newKbSysId,
    );

    return { kbSysId: newKbSysId, domain: domain };
  },

  // Dual-mode function:
  //   checkCompany(request)            — bulk check from request body; returns { companyNames: [bool, ...] }
  //   checkCompany(null, companyName)  — single lookup; returns the sys_domain value of the matching core_company, or null
  checkCompany: function (request, companyName) {
    if (!request && companyName) {
      var gr = new GlideRecord("core_company");
      gr.addQuery("name", companyName);
      gr.setLimit(1);
      gr.query();
      return gr.next() ? gr.getValue("sys_domain") : null;
    }

    this._logSeparator("Initialize checkCompany");

    var body = request.body.data;

    if (
      !body ||
      !Array.isArray(body.companyNames) ||
      body.companyNames.length === 0
    ) {
      this._log(
        "warn",
        "Request received with missing or empty companyNames array.",
      );
      return { companyNames: [] };
    }

    var companyNames = body.companyNames;
    this._log("info", "Bulk company check — count: " + companyNames.length);

    var results = [];

    for (var i = 0; i < companyNames.length; i++) {
      var name = String(companyNames[i]).trim();
      var grCheck = new GlideRecord("core_company");
      grCheck.addQuery("name", name);
      grCheck.setLimit(1);
      grCheck.query();

      var found = grCheck.next();
      this._log(
        found ? "info" : "warn",
        "  [" +
          (i + 1) +
          "/" +
          companyNames.length +
          '] "' +
          name +
          '" — ' +
          (found ? "FOUND" : "NOT FOUND"),
      );
      results.push(found);
    }

    this._logSeparator("Finalize checkCompany");
    return { companyNames: results };
  },

  type: "ITGlueImportDocuments",
};
