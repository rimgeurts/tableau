let app = new Vue({
  el: "#app",
  data: {
    fileName: null,
    errorMessage: null,
    sheetMap: [],
    fieldMap: [],
    worksheetSources: [],
    currentTab: 1,
    showSheets: true,
  },
  methods: {
    stripBrackets: function (string) {
      return (string = string.startsWith("[") && string.endsWith("]") ? string.slice(1, -1) : string);
    },
    processFile: function (event) {
      this.errorMessage = null;
      let workbook = event.target.files[0];
      this.fileName = workbook.name;
      let type = workbook.name.split(".").slice(-1)[0];

      if (type === "twbx") {
        let zip = new JSZip();
        zip.loadAsync(workbook).then(
          (zip) => {
            const twbName = Object.keys(zip.files).find((file) => file.endsWith(".twb"));
            const twb = zip.files[twbName];
            twb.async("string").then((content) => {
              if (!content) return (this.errorMessage = "No twb file found!");
              this.parseXML(content);
            });
          },
          () => {
            alert("Not a valid twbx file");
          }
        );
      } else if (type === "twb") {
        let reader = new FileReader();
        reader.onload = (evt) => {
          if (!evt.target.result) return (this.errorMessage = "No file found!");
          this.parseXML(evt.target.result);
        };
        reader.readAsText(workbook);
      } else {
        this.errorMessage = "File was not a twb or twbx.";
      }
    },
    parseXML: function (text) {
      this.sheetMap = [];
      this.fieldMap = [];
      let parser = new DOMParser();
      let xml = parser.parseFromString(text, "text/xml");
      this.getSheets(xml);
      this.getFields(xml);
      this.addCalcDef(xml);
    },
    getSheets: function (xml) {
      let sheetMap = [];
      let worksheetSources = [];
      if (xml.getElementsByTagName("worksheets").length > 0) {
        let worksheets = xml.getElementsByTagName("worksheets")[0].children;
        for (let worksheet of worksheets) {
          let wsName = worksheet.attributes.name.nodeValue;
          worksheetSources.push({ wsName, dataSources: [] });
          let dataSources = worksheet.getElementsByTagName("table")[0].getElementsByTagName("view")[0].getElementsByTagName("datasources")[0].children;
          for (let dataSource of dataSources) {
            let dsName = dataSource.attributes.caption ? dataSource.attributes.caption.nodeValue : dataSource.attributes.name.nodeValue;
            let dsID = dataSource.attributes.name.nodeValue;
            let foundDS = sheetMap.find((d) => d.dsName == dsName);
            if (foundDS) {
              foundDS.sheets.push({ name: wsName, type: "worksheet" });
            } else {
              sheetMap.push({
                dsName,
                sheets: [{ name: wsName, type: "worksheet" }],
              });
            }
            let foundWS = worksheetSources.find((ws) => ws.wsName === wsName);
            foundWS.dataSources.push({ name: dsName, id: dsID });
          }
        }
      }
      this.worksheetSources = worksheetSources;
      if (xml.getElementsByTagName("dashboards").length > 0) {
        let dashboards = xml.getElementsByTagName("dashboards")[0].children;
        for (let dashboard of dashboards) {
          let dbName = dashboard.attributes.name.nodeValue;
          if (dashboard.getElementsByTagName("datasources").length > 0) {
            let dataSources = dashboard.getElementsByTagName("datasources")[0].children;
            for (let dataSource of dataSources) {
              let dsName = dataSource.attributes.caption ? dataSource.attributes.caption.nodeValue : dataSource.attributes.name.nodeValue;
              let foundDS = sheetMap.find((d) => d.dsName == dsName);
              if (foundDS) {
                foundDS.sheets.push({ name: dbName, type: "dashboard" });
              } else {
                sheetMap.push({
                  dsName,
                  sheets: [{ name: dbName, type: "dashboard" }],
                });
              }
            }
          }
          if (dashboard.getElementsByTagName("zones").length > 0) {
            let zones = dashboard.getElementsByTagName("zone");
            for (let zone of zones) {
              if (zone.attributes.name && zone.attributes.id) {
                let wsName = zone.attributes.name.nodeValue;
                let foundWS = worksheetSources.find((ws) => ws.wsName === wsName);
                if (foundWS) {
                  for (let source of foundWS.dataSources) {
                    let foundDS = sheetMap.find((ds) => ds.dsName === source.name);
                    if (!foundDS.sheets.find((sheet) => sheet.name === dbName)) {
                      foundDS.sheets.push({ name: dbName, type: "dashboard" });
                    }
                  }
                }
              }
            }
          }
        }
      }
      this.sheetMap = sheetMap;
      if (Object.keys(sheetMap).length === 0) this.errorMessage = "No worksheets or dashboards found.";
    },
    getFields: function (xml) {
      let fieldMap = [];
      let calcDef = [];
      if (xml.getElementsByTagName("worksheets").length > 0) {
        let worksheets = xml.getElementsByTagName("worksheets")[0].children;
        for (let worksheet of worksheets) {
          let wsName = worksheet.attributes.name.nodeValue;
          let dsDependencies = worksheet.getElementsByTagName("datasource-dependencies");
          for (let dataSource of dsDependencies) {
            let dsID = dataSource.attributes.datasource.nodeValue;
            let foundDS = this.worksheetSources.find((ws) => ws.wsName === wsName).dataSources.find((ds) => ds.id === dsID);
            if (!foundDS) continue;
            let dsName = foundDS.name;
            if (!fieldMap.find((ds) => ds.dsName === dsName)) fieldMap.push({ dsName, fields: [] });
            let dsFields = fieldMap.find((ds) => ds.dsName === dsName);
            let columns = dataSource.getElementsByTagName("column");
            for (let column of columns) {
              let fieldName = column.attributes.caption ? column.attributes.caption.nodeValue : column.attributes.name.nodeValue;
              fieldName = this.stripBrackets(fieldName);
              let type = column.getElementsByTagName("calculation").length > 0 ? "calculation" : "datasourcefield";
              let calc =
                type === "calculation" && column.getElementsByTagName("calculation")[0].attributes.formula
                  ? column.getElementsByTagName("calculation")[0].attributes.formula.nodeValue
                  : null;
              let foundField = dsFields.fields.find((f) => f.name === fieldName);
              if (foundField) {
                foundField.worksheets.push(wsName);
              } else {
                dsFields.fields.push({
                  name: fieldName,
                  type,
                  calc,
                  worksheets: [wsName],
                });
              }
            }
          }
        }
      }
      for (let ds of fieldMap) {
        ds = ds.fields.sort((a, b) => (a.name > b.name ? 1 : -1));
      }
      this.fieldMap = fieldMap;
    },
    addCalcDef: function (xml) {
      let calcDef = [];
      let calcList = [];
      if (xml.getElementsByTagName("datasources").length > 0) {
        let datasources = xml.getElementsByTagName("datasources")[0].children;
        for (let dataSource of datasources) {
          let dsName = dataSource.attributes.caption ? dataSource.attributes.caption.nodeValue : dataSource.attributes.name.nodeValue;
          let dsID = dataSource.attributes.name.nodeValue;
          calcDef.push({ dsName, dsID, columns: [] });
          let ds = calcDef.find((ds) => ds.dsName === dsName);
          let columns = dataSource.getElementsByTagName("column");
          for (let column of columns) {
            let isCalc = column.getElementsByTagName("calculation").length > 0;
            if (isCalc) {
              let name = column.attributes.caption ? column.attributes.caption.nodeValue : column.attributes.name.nodeValue;
              let id = column.attributes.name.nodeValue;
              ds.columns.push({ name, id });
            }
          }
        }
      }
      for (let ds of calcDef) {
        let dsName = ds.dsName;
        for (let calc of ds.columns) {
          let name = this.stripBrackets(calc.name);
          let id = calc.id;
          let displayDS = this.fieldMap.find((ds) => ds.dsName === dsName);
          if (!displayDS) continue;
          for (let field of displayDS.fields) {
            if (field.calc) {
              field.calc = field.calc.replaceAll(id, `[${name}]`);
              for (let ds2 of calcDef) {
                let ds2Name = this.stripBrackets(ds2.dsName);
                let ds2ID = ds2.dsID;
                if (ds2ID !== "Parameters") field.calc = field.calc.replaceAll(ds2ID, `[${ds2Name}]`);
              }
              if (!calcList.find((f) => f.dsName === dsName && f.name === field.name)) calcList.push({ dsName, name: field.name, calc: field.calc });
            }
          }
        }
      }
      for (let calc of calcList) {
        if (calc.dsName !== "Parameters") {
          let name = calc.name;
          let r = new RegExp(/\[([^\[\]]+)\]/g);
          let matches = calc.calc.match(r);
          if (matches && matches.length > 0) {
            for (let match of matches) {
              let depName = this.stripBrackets(match);
              for (let ds of this.fieldMap) {
                for (let field of ds.fields) {
                  if (depName === field.name && field.worksheets.indexOf(`=${name}`) === -1) field.worksheets.push(`=${name}`);
                }
              }
            }
          }
        }
      }
    },
  },
});
