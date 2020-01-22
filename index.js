// node-gdoctableapp: This is a Node.js module to manage the tables on Google Document using Google Docs API.
const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const version = "1.1.0";

function getValuesFromTable(content) {
  return content.map(function(row) {
    return row.reduce(function(ar, e) {
      const temp = e.map(function(f) {
        return f.content.replace("\n", "");
      });
      Array.prototype.push.apply(ar, temp);
      return ar;
    }, []);
  });
}

function getTablesMain(obj, callback) {
  const tables = obj.docTables;
  const res = tables.map(function(table, i) {
    obj.docTable = table;
    parseTable(obj);
    const values = getValuesFromTable(obj.content);
    return {
      index: i,
      values: values,
      tablePosition: {
        startIndex: table.startIndex,
        endIndex: table.endIndex
      }
    };
  });
  obj.result.tables = res;
  callback(null, obj);
  return;
}

function getValuesMain(obj, callback) {
  parseTable(obj);
  const values = getValuesFromTable(obj.content);
  obj.result.values = values;
  callback(null, obj);
  return;
}

function deleteTableMain(obj, callback) {
  obj.requestBody = [
    {
      deleteContentRange: {
        range: {
          startIndex: obj.docTable.startIndex,
          endIndex: obj.docTable.endIndex
        }
      }
    }
  ];
  documentsBatchUpdate(obj)
    .then(function(res) {
      obj.result.responseFromAPIs.push(res.data);
      callback(null, obj);
    })
    .catch(function(err) {
      callback(err, null);
    });
}

function deleteRowsAndColumnsMain(obj, callback) {
  const maxDeleteRows = Math.max.apply(null, obj.params.deleteRows) + 1;
  const maxDeleteCols = Math.max.apply(null, obj.params.deleteColumns) + 1;
  const table = obj.docTable.table;
  if (table.rows < maxDeleteRows || table.columns < maxDeleteCols) {
    return callback(
      { errors: ["Rows and columns for deleting are outside of the table."] },
      null
    );
  }
  const tablePos = obj.docTable.startIndex;
  let iObj = obj.params;
  let requests = [];

  if (
    "deleteRows" in iObj &&
    Array.isArray(iObj.deleteRows) &&
    iObj.deleteRows.length > 0
  ) {
    iObj.deleteRows = descendingSort(iObj.deleteRows);
    for (let i = 0; i < iObj.deleteRows.length; i++) {
      requests.push({
        deleteTableRow: {
          tableCellLocation: {
            tableStartLocation: { index: tablePos },
            rowIndex: iObj.deleteRows[i]
          }
        }
      });
    }
  }
  if (
    "deleteColumns" in iObj &&
    Array.isArray(iObj.deleteColumns) &&
    iObj.deleteColumns.length > 0
  ) {
    iObj.deleteColumns = descendingSort(iObj.deleteColumns);
    for (let i = 0; i < iObj.deleteColumns.length; i++) {
      requests.push({
        deleteTableColumn: {
          tableCellLocation: {
            tableStartLocation: { index: tablePos },
            columnIndex: iObj.deleteColumns[i]
          }
        }
      });
    }
  }
  obj.requestBody = requests;
  documentsBatchUpdate(obj)
    .then(function(res) {
      obj.result.responseFromAPIs.push(res.data);
      callback(null, obj);
    })
    .catch(function(err) {
      callback(err, null);
    });
}

function setValuesMain(obj, callback) {
  if (!("values" in obj.params)) {
    return callback({ errors: ["Please set 'values'."] }, null);
  }
  if (valuesChecker(obj)) {
    obj.params.values = [
      {
        values: obj.params.values,
        range: { startRowIndex: 0, startColumnIndex: 0 }
      }
    ];
  }
  var dupChk = checkDupValues(obj);
  if (dupChk.dup.length > 0) {
    return callback(
      { errors: ["Range of inputted values are duplicated."] },
      null
    );
  }
  obj.requests = {};
  parseInputValuesForSetValues(obj, dupChk);
  addRowsAndColumnsForSetValues(obj);

  Promise.resolve()
    .then(function() {
      return new Promise(function(resolve, reject) {
        addRowsAndColumnsByAPI(obj, function(err, obj) {
          if (err) {
            reject(err);
            return;
          }
          resolve(obj);
        });
      });
    })
    .then(function(obj) {
      return new Promise(function(resolve) {
        parseTable(obj);
        resolve(obj);
      });
    })
    .then(function(obj) {
      return new Promise(function(resolve, reject) {
        obj.requestBody = createRequestsForSetValues(obj);
        documentsBatchUpdate(obj)
          .then(function(res) {
            obj.result.responseFromAPIs.push(res.data);
            resolve(obj);
          })
          .catch(function(err) {
            reject(err);
          });
      });
    })
    .then(function(obj) {
      callback(null, obj);
    })
    .catch(function(err) {
      callback(err, null);
    });
}

function insertTable(obj, callback) {
  let requests = [];
  requests.push({
    insertTable: {
      rows: obj.params.rows,
      columns: obj.params.columns,
      location: {
        index: obj.params.createIndex
      }
    }
  });
  if ("values" in obj.params && obj.params.values.length > 0) {
    createRequestBodyForInsertText(obj, requests, obj.params.createIndex);
  }
  obj.requestBody = requests;
  documentsBatchUpdate(obj)
    .then(function(res) {
      obj.result.responseFromAPIs.push(res.data);
      callback(null, obj);
    })
    .catch(function(err) {
      callback(err, null);
    });
}

function appendTable(obj, callback) {
  Promise.resolve()
    .then(function() {
      return new Promise(function(resolve, reject) {
        obj.requestBody = [
          {
            insertTable: {
              rows: obj.params.rows,
              columns: obj.params.columns,
              endOfSegmentLocation: { segmentId: "" }
            }
          }
        ];
        documentsBatchUpdate(obj)
          .then(function(res) {
            obj.result.responseFromAPIs.push(res.data);
            resolve(obj);
          })
          .catch(function(err) {
            reject(err);
          });
      });
    })
    .then(function(obj) {
      return new Promise(function(resolve, reject) {
        if ("values" in obj.params && obj.params.values.length > 0) {
          getDocument(obj)
            .then(function(contents) {
              let table = {};
              for (let i = contents.length - 1; i >= 0; i--) {
                const content = contents[i];
                if (content.table) {
                  table = content;
                  break;
                }
              }
              obj.docTable = table;
              obj.result.responseFromAPIs.push(contents);
              resolve(obj);
            })
            .catch(function(err) {
              reject(err);
            });
        } else {
          resolve(obj);
        }
      });
    })
    .then(function(obj) {
      return new Promise(function(resolve, reject) {
        if ("values" in obj.params && obj.params.values.length > 0) {
          let requests = [];
          createRequestBodyForInsertText(
            obj,
            requests,
            obj.docTable.startIndex - 1
          );
          obj.requestBody = requests;
          documentsBatchUpdate(obj)
            .then(function(res) {
              obj.result.responseFromAPIs.push(res.data);
              resolve(obj);
            })
            .catch(function(err) {
              reject(err);
            });
        }
      });
    })
    .then(function(obj) {
      callback(null, obj);
    })
    .catch(function(err) {
      callback(err, null);
    });
}

function createTableMain(obj, callback) {
  if (!("createIndex" in obj.params) && !("append" in obj.params)) {
    callback("Please set 'createIndex' or 'append'.", null);
    return;
  }
  if (!("rows" in obj.params) || !("columns" in obj.params)) {
    callback(
      { errors: ["Please set rows and columns for creating new table."] },
      null
    );
    return;
  }
  if (obj.params.append) {
    appendTable(obj, function(err, res) {
      if (err) {
        callback(err, null);
        return;
      }
      callback(null, res);
    });
  } else if (obj.params.createIndex) {
    insertTable(obj, function(err, res) {
      if (err) {
        callback(err, null);
        return;
      }
      callback(null, res);
    });
  } else {
    callback({ errors: ["Please set Index (> 0) or Append."] }, null);
  }
}

function appendRowMain(obj, callback) {
  if (!("values" in obj.params) || obj.params.values.length == 0) {
    callback({ errors: ["Values for putting are not set."] }, null);
  }
  obj.params.values = [
    {
      values: obj.params.values,
      range: { startRowIndex: obj.docTable.table.rows, startColumnIndex: 0 }
    }
  ];
  setValuesMain(obj, function(err, obj) {
    if (err) {
      callback(err, null);
      return;
    }
    callback(null, obj);
  });
}

function getTextRunContent(obj, ar, h) {
  if ("paragraph" in h) {
    for (let i = 0; i < h.paragraph.elements.length; i++) {
      const e = h.paragraph.elements[i];
      if (
        "textRun" in e &&
        e.textRun.content.indexOf(obj.params.searchText) != -1
      ) {
        ar.push(e);
      }
    }
  }
}

function getTableContent(obj, e) {
  let ar = [];
  for (let i = 0; i < e.table.tableRows.length; i++) {
    const f = e.table.tableRows[i];
    for (let j = 0; j < f.tableCells.length; j++) {
      const g = f.tableCells[j];
      for (let k = 0; k < g.content.length; k++) {
        const h = g.content[k];
        getTextRunContent(obj, ar, h);
      }
    }
  }
  return ar;
}

function createInsertInlineImageRequest(startIndex, url, width, height) {
  let req = {
    insertInlineImage: {
      uri: url,
      location: { index: startIndex }
    }
  };
  if (!isNaN(width) && !isNaN(height) && width > 0 && height > 0) {
    req.insertInlineImage.objectSize = {
      width: { magnitude: width, unit: "PT" },
      height: { magnitude: height, unit: "PT" }
    };
  }
  return req;
}

function deleteTempFile(obj, callback) {
  if (
    "replaceImageFilePath" in obj.params &&
    obj.params.replaceImageFilePath != "" &&
    obj.tempFileId != ""
  ) {
    const drive = google.drive({
      version: "v3",
      auth: obj.params.auth
    });
    drive.files.delete({ fileId: obj.tempFileId }, function(err) {
      if (err) {
        callback(err, null);
        return;
      }
      callback(null, "done");
    });
  } else {
    callback(null, "done");
  }
}

function uploadImageFile(obj, callback) {
  if (
    (!("replaceImageURL" in obj.params) || !obj.params.replaceImageURL) &&
    "replaceImageFilePath" in obj.params &&
    obj.params.replaceImageFilePath != ""
  ) {
    const drive = google.drive({ version: "v3", auth: obj.params.auth });
    drive.files.create(
      {
        requestBody: {
          name: path.basename(obj.params.replaceImageFilePath)
        },
        media: {
          body: fs.createReadStream(obj.params.replaceImageFilePath)
        },
        fields: "id,webContentLink"
      },
      function(err, file) {
        if (err) {
          callback(err, null);
          return;
        }
        obj.result.responseFromAPIs.push(file.data);
        obj.params.replaceImageURL = file.data.webContentLink;
        obj.tempFileId = file.data.id;
        drive.permissions.create(
          {
            fileId: file.data.id,
            requestBody: {
              type: "anyone",
              role: "reader"
            }
          },
          function(err, file) {
            if (err) {
              callback(err, null);
              return;
            }
            obj.result.responseFromAPIs.push(file.data);
            callback(null, "done");
          }
        );
      }
    );
  } else {
    callback(null, "done");
  }
}

function replaceTextsToImagesByURL(obj, callback) {
  const contents = obj.docTables.reduce(function(a, e) {
    if ("table" in e) {
      Array.prototype.push.apply(a, getTableContent(obj, e));
    } else if ("paragraph" in e) {
      getTextRunContent(obj, a, e);
    }
    return a;
  }, []);
  const searchText = obj.params.searchText;
  const replacedUrl = obj.params.replaceImageURL;
  const width = obj.params.imageWidth;
  const height = obj.params.imageHeight;
  const requests = contents.reverse().reduce(function(ar, e) {
    const content = e.textRun.content;
    if (content.trim() == searchText) {
      const offset = content.length - content.trim().length;
      ar.push({
        deleteContentRange: {
          range: { startIndex: e.startIndex, endIndex: e.endIndex - offset }
        }
      });
      ar.push(
        createInsertInlineImageRequest(e.startIndex, replacedUrl, width, height)
      );
    } else {
      const start = e.startIndex + content.indexOf(searchText);
      ar.push({
        deleteContentRange: {
          range: {
            startIndex: start,
            endIndex: start + searchText.length
          }
        }
      });
      ar.push(
        createInsertInlineImageRequest(start, replacedUrl, width, height)
      );
    }
    return ar;
  }, []);
  if (requests.length > 0) {
    obj.requestBody = requests;
    documentsBatchUpdate(obj)
      .then(function(res) {
        obj.result.responseFromAPIs.push(res.data);
        callback(null, "done");
      })
      .catch(function(err) {
        callback(err, null);
      });
  } else {
    callback(`'${searchText}' was not found.`, null);
  }
}

function replaceTextsToImagesMain(obj, callback) {
  Promise.resolve()
    .then(function() {
      return new Promise(function(resolve, reject) {
        uploadImageFile(obj, function(err, res) {
          if (err) {
            reject(err);
            return;
          }
          resolve(res);
        });
      });
    })
    .then(function() {
      return new Promise(function(resolve, reject) {
        replaceTextsToImagesByURL(obj, function(err, res) {
          if (err) {
            reject(err);
            return;
          }
          resolve(res);
        });
      });
    })
    .then(function() {
      return new Promise(function(resolve, reject) {
        deleteTempFile(obj, function(err, res) {
          if (err) {
            reject(err);
            return;
          }
          resolve(res);
        });
      });
    })
    .then(function() {
      callback(null, obj);
    })
    .catch(function(err) {
      callback(err, null);
    });
}

function createRequestBodyForInsertText(obj, requests, idx) {
  const val = parseInputValues(
    obj.params.values,
    idx,
    obj.params.rows,
    obj.params.columns
  );
  for (let i = val.length - 1; i >= 0; i--) {
    const v = val[i].content;
    if (v != "") {
      requests.push({
        insertText: {
          location: {
            index: val[i].index
          },
          text: v
        }
      });
    }
  }
}

function parseInputValues(values, index, rows, cols) {
  index += 4;
  const v = [];
  let maxCol;
  const maxRow = values.length;
  for (let row = 0; row < rows; row++) {
    if (maxRow > row) {
      maxCol = values[row].length;
    } else {
      maxCol = cols;
    }
    for (let col = 0; col < cols; col++) {
      if (maxRow > row && maxCol > col && values[row][col] != "") {
        v.push({
          row: row,
          col: col,
          content: values[row][col],
          index: index
        });
      }
      index += 2;
    }
    index++;
  }
  return v;
}

function addRowsAndColumnsByAPI(obj, callback) {
  const tr = obj.requests.insertTableRow;
  const tc = obj.requests.insertTableColumn;
  if (tr.length > 0 || tc.length > 0) {
    let requests = [];
    if (tr.length > 0) {
      for (let i = 0; i < tr.length; i++) {
        requests.push(tr[i]);
      }
    }
    if (tc.length > 0) {
      for (let i = 0; i < tc.length; i++) {
        requests.push(tc[i]);
      }
    }
    obj.requestBody = requests;

    Promise.resolve()
      .then(function() {
        return new Promise(function(resolve, reject) {
          documentsBatchUpdate(obj)
            .then(function(res) {
              obj.result.responseFromAPIs.push(res.data);
              resolve(obj);
            })
            .catch(function(err) {
              reject(err);
            });
        });
      })
      .then(function(obj) {
        return new Promise(function(resolve, reject) {
          obj.requestBody = {};
          getTable(obj, function(err, obj) {
            if (err) {
              reject(err);
              return;
            }
            resolve(obj);
          });
        });
      })
      .then(function(obj) {
        callback(null, obj);
      })
      .catch(function(err) {
        callback(err, null);
      });
    delete obj.requests.insertTableColumn;
    delete obj.requests.insertTableRow;
  } else {
    callback(null, obj);
  }
}

function createRequestsForSetValues(obj) {
  let requests = [];
  const values = obj.parsedValues;
  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i].row;
    const c = values[i].col;
    const v = values[i].content.toString();
    const delReq = obj.delCell[r][c];
    if (
      delReq.deleteContentRange.range.startIndex !=
      delReq.deleteContentRange.range.endIndex
    ) {
      requests.push(delReq);
    }
    if (v != "") {
      requests.push({
        insertText: {
          location: { index: delReq.deleteContentRange.range.startIndex },
          text: v
        }
      });
    }
  }
  return requests;
}

function addRowsAndColumns(startIndex, maxRow, maxCol, tableRow, tableCol) {
  const addRows = maxRow - tableRow;
  const addColumns = maxCol - tableCol;
  let obj = { insertTableRowBody: [], insertTableColumnBody: [] };
  if (addRows > 0) {
    for (let i = 0; i < addRows; i++) {
      obj.insertTableRowBody.push({
        insertTableRow: {
          insertBelow: true,
          tableCellLocation: {
            tableStartLocation: { index: startIndex },
            rowIndex: tableRow - 1 + i
          }
        }
      });
    }
  }
  if (addColumns > 0) {
    for (let i = 0; i < addColumns; i++) {
      obj.insertTableColumnBody.push({
        insertTableColumn: {
          insertRight: true,
          tableCellLocation: {
            tableStartLocation: { index: startIndex },
            columnIndex: tableCol - 1 + i
          }
        }
      });
    }
  }
  return obj;
}

function addRowsAndColumnsForSetValues(obj) {
  const values = obj.params.values;
  const res = values.reduce(
    function(o, e) {
      const maxRow = e.values.length + e.range.startRowIndex;
      const maxCol =
        e.values.reduce(function(n, f) {
          if (n < f.length) n = f.length;
          return n;
        }, 0) + e.range.startColumnIndex;
      if (o.maxRow < maxRow) o.maxRow = maxRow;
      if (o.maxCol < maxCol) o.maxCol = maxCol;
      return o;
    },
    { maxRow: 0, maxCol: 0 }
  );
  const o = addRowsAndColumns(
    obj.docTable.startIndex,
    res.maxRow,
    res.maxCol,
    obj.docTable.table.rows,
    obj.docTable.table.columns
  );
  obj.requests.insertTableRow = o.insertTableRowBody;
  obj.requests.insertTableColumn = o.insertTableColumnBody;
}

function parseInputValuesForSetValues(obj, dupChk) {
  dupChk.noDup.sort(function(a, b) {
    if (a.col < b.col) return -1;
    if (a.col > b.col) return 1;
    return 0;
  });
  dupChk.noDup.sort(function(a, b) {
    if (a.row < b.row) return -1;
    if (a.row > b.row) return 1;
    return 0;
  });
  obj.parsedValues = dupChk.noDup;
}

function checkDupValues(obj) {
  const values = obj.params.values;
  const temp = values.reduce(function(ar1, e) {
    const rowOffset = e.range.startRowIndex;
    const colOffset = e.range.startColumnIndex;
    const temp1 = e.values.reduce(function(ar2, row, i) {
      const temp2 = row.map(function(col, j) {
        return { row: i + rowOffset, col: j + colOffset, content: col };
      });
      Array.prototype.push.apply(ar2, temp2);
      return ar2;
    }, []);

    Array.prototype.push.apply(ar1, temp1);
    return ar1;
  }, []);

  const dupCheck = temp.reduce(
    function(o, e) {
      if (
        o.noDup.some(function(f) {
          return f.row === e.row && f.col === e.col;
        })
      ) {
        o.dup.push(e);
      } else {
        o.noDup.push(e);
      }
      return o;
    },
    { dup: [], noDup: [] }
  );

  return dupCheck;
}

function valuesChecker(obj) {
  return obj.params.values.every(function(e) {
    return Array.isArray(e) && !e.values && !e.range;
  });
}

function descendingSort(ar) {
  return ar.sort(function(a, b) {
    if (a > b) return -1;
    if (a < b) return 1;
    return 0;
  });
}

function parseTable(obj) {
  const docContent = obj.docTable;
  const tableRows = docContent.table.tableRows;
  let valuesIndexes = { deleteIndex: [], content: [] };
  for (let i = 0; i < tableRows.length; i++) {
    const tableCells = tableRows[i].tableCells;
    let tempRowsDelCell = [];
    let tempRowsContent = [];
    for (let j = 0; j < tableCells.length; j++) {
      let tempColsDelCell = { deleteContentRange: { range: {} } };
      let tempColsContent = [];
      const contents = tableCells[j].content;
      for (let k = 0; k < contents.length; k++) {
        if ("paragraph" in contents[k]) {
          const elements = contents[k].paragraph.elements;
          for (var l = 0; l < elements.length; l++) {
            if (k == 0 && l == 0) {
              tempColsDelCell.deleteContentRange.range.startIndex =
                elements[l].startIndex;
            }
            if (k == contents.length - 1 && l == elements.length - 1) {
              tempColsDelCell.deleteContentRange.range.endIndex =
                elements[l].endIndex - 1;
            }
            let cellContent = "";
            if ("textRun" in elements[l]) {
              cellContent = elements[l].textRun.content;
            } else if ("inlineObjectElement" in elements[l]) {
              cellContent = "[INLINE OBJECT]";
            } else {
              cellContent = "[UNSUPPORTED CONTENT]";
            }
            tempColsContent.push({
              startIndex: elements[l].startIndex,
              endIndex: elements[l].endIndex,
              content: cellContent
            });
          }
        } else if ("table" in contents[k]) {
          tempColsContent.push({
            startIndex: contents[k].startIndex,
            endIndex: contents[k].endIndex,
            content: "[TABLE]"
          });
        } else {
          tempColsContent.push({
            startIndex: contents[k].startIndex,
            endIndex: contents[k].endIndex,
            content: "[UNSUPPORTED CONTENT]"
          });
        }
      }

      tempRowsDelCell.push(tempColsDelCell);
      tempRowsContent.push(tempColsContent);
    }
    valuesIndexes.deleteIndex.push(tempRowsDelCell);
    valuesIndexes.content.push(tempRowsContent);
  }
  obj.delCell = valuesIndexes.deleteIndex;
  obj.content = valuesIndexes.content;
  obj.cell1stIndex = valuesIndexes.content[0][0][0].startIndex;
}

function getAllContents(obj, callback) {
  getDocument(obj)
    .then(function(contents) {
      obj.docTables = contents;
      obj.result.responseFromAPIs.push(contents);
      callback(null, obj);
    })
    .catch(function(err) {
      callback(err, null);
    });
}

function getAllTables(obj, callback) {
  getDocument(obj)
    .then(function(contents) {
      let tables = [];
      for (let i = 0; i < contents.length; i++) {
        const content = contents[i];
        if (content.table) {
          tables.push(content);
        }
      }
      obj.docTables = tables;
      obj.result.responseFromAPIs.push(contents);
      callback(null, obj);
    })
    .catch(function(err) {
      callback(err, null);
    });
}

function getTable(obj, callback) {
  getDocument(obj)
    .then(function(contents) {
      let ti = 0;
      let table = {};
      for (let i = 0; i < contents.length; i++) {
        const content = contents[i];
        if (content.table) {
          if (ti == obj.params.tableIndex) {
            table = content;
            break;
          }
          ti++;
        }
      }
      if (Object.keys(table).length == 0) {
        callback(
          {
            errors: [`No table is found at index '${obj.params.tableIndex}'.`]
          },
          null
        );
        return;
      }
      obj.docTable = table;
      obj.result.responseFromAPIs.push(contents);
      callback(null, obj);
    })
    .catch(function(err) {
      callback(err, null);
    });
}

async function documentsBatchUpdate(obj) {
  return await obj.srv.documents.batchUpdate({
    documentId: obj.params.documentId,
    resource: { requests: obj.requestBody }
  });
}

async function getDocument(obj) {
  const document = await obj.srv.documents.get({
    documentId: obj.params.documentId
  });
  return document.data.body.content;
}

function checkAuth(auth) {
  if (auth instanceof Object) {
    if ("credentials" in auth && "access_token" in auth.credentials) {
      return true;
    } else if ("key" in auth && "email" in auth) {
      return true;
    }
  }
  return false;
}

function init(e, callback) {
  if (!("documentId" in e)) {
    callback({ errors: ["Please set 'documentId'."] }, null);
    return;
  }
  const chkAuth = checkAuth(e.auth);
  if (!chkAuth) {
    callback({ errors: ["Please use OAuth2 or Service account."] }, null);
    return;
  }
  let obj = {
    params: e,
    srv: google.docs({ version: "v1", auth: e.auth }),
    result: { responseFromAPIs: [], libraryVersion: version }
  };
  if ("tableIndex" in obj.params) {
    if (obj.params.tableIndex == -1) {
      getAllTables(obj, function(err, obj) {
        if (err) {
          callback(err, null);
          return;
        }
        callback(null, obj);
      });
    } else if (obj.params.tableIndex == -2) {
      getAllContents(obj, function(err, obj) {
        if (err) {
          callback(err, null);
          return;
        }
        callback(null, obj);
      });
    } else {
      getTable(obj, function(err, obj) {
        if (err) {
          callback(err, null);
          return;
        }
        callback(null, obj);
      });
    }
  } else {
    callback(null, obj);
  }
}

function getTables(params, callback) {
  params.tableIndex = -1;
  init(params, function(err, obj) {
    if (err) {
      callback(err, null);
      return;
    }
    getTablesMain(obj, function(err, obj) {
      if (err) {
        callback(err, null);
        return;
      }
      if (!obj.params.showAPIResponse) delete obj.result.responseFromAPIs;
      callback(null, obj.result);
    });
  });
}

function getValues(params, callback) {
  if (!("tableIndex" in params)) {
    callback({ errors: ["Please set 'tableIndex'."] }, null);
    return;
  }
  init(params, function(err, obj) {
    if (err) {
      callback(err, null);
      return;
    }
    getValuesMain(obj, function(err, obj) {
      if (err) {
        callback(err, null);
        return;
      }
      if (!obj.params.showAPIResponse) delete obj.result.responseFromAPIs;
      callback(null, obj.result);
    });
  });
}

function deleteTable(params, callback) {
  if (!("tableIndex" in params)) {
    callback({ errors: ["Please set 'tableIndex'."] }, null);
    return;
  }
  init(params, function(err, obj) {
    if (err) {
      callback(err, null);
      return;
    }
    deleteTableMain(obj, function(err, obj) {
      if (err) {
        callback(err, null);
        return;
      }
      if (!obj.params.showAPIResponse) delete obj.result.responseFromAPIs;
      callback(null, obj.result);
    });
  });
}

function deleteRowsAndColumns(params, callback) {
  if (!("tableIndex" in params)) {
    callback({ errors: ["Please set 'tableIndex'."] }, null);
    return;
  }
  init(params, function(err, obj) {
    if (err) {
      callback(err, null);
      return;
    }
    deleteRowsAndColumnsMain(obj, function(err, obj) {
      if (err) {
        callback(err, null);
        return;
      }
      if (!obj.params.showAPIResponse) delete obj.result.responseFromAPIs;
      callback(null, obj.result);
    });
  });
}

function setValues(params, callback) {
  if (!("tableIndex" in params)) {
    callback({ errors: ["Please set 'tableIndex'."] }, null);
    return;
  }
  init(params, function(err, obj) {
    if (err) {
      callback(err, null);
      return;
    }
    setValuesMain(obj, function(err, obj) {
      if (err) {
        callback(err, null);
        return;
      }
      if (!obj.params.showAPIResponse) delete obj.result.responseFromAPIs;
      callback(null, obj.result);
    });
  });
}

function createTable(params, callback) {
  init(params, function(err, obj) {
    if (err) {
      callback(err, null);
      return;
    }
    createTableMain(obj, function(err, obj) {
      if (err) {
        callback(err, null);
        return;
      }
      if (!obj.params.showAPIResponse) delete obj.result.responseFromAPIs;
      callback(null, obj.result);
    });
  });
}

function appendRow(params, callback) {
  if (!("tableIndex" in params)) {
    callback({ errors: ["Please set 'tableIndex'."] }, null);
    return;
  }
  init(params, function(err, obj) {
    if (err) {
      callback(err, null);
      return;
    }
    appendRowMain(obj, function(err, obj) {
      if (err) {
        callback(err, null);
        return;
      }
      if (!obj.params.showAPIResponse) delete obj.result.responseFromAPIs;
      callback(null, obj.result);
    });
  });
}

function replaceTextsToImages(params, callback) {
  params.tableIndex = params.tableOnly ? -1 : -2;
  init(params, function(err, obj) {
    if (err) {
      callback(err, null);
      return;
    }
    replaceTextsToImagesMain(obj, function(err, obj) {
      if (err) {
        callback(err, null);
        return;
      }
      if (!obj.params.showAPIResponse) delete obj.result.responseFromAPIs;
      callback(null, obj.result);
    });
  });
}

module.exports = {
  GetTables: getTables,
  GetValues: getValues,
  SetValues: setValues,
  DeleteTable: deleteTable,
  DeleteRowsAndColumns: deleteRowsAndColumns,
  CreateTable: createTable,
  AppendRow: appendRow,
  ReplaceTextsToImages: replaceTextsToImages
};
