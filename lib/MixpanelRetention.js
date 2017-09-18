'use strict';

var MixpanelExport = require('mixpanel-data-export-node');
var Moment = require('moment');
var XLSX = require('xlsx-style');
var _ = require('lodash');
var Jszip = require('jszip');
var Fs = require('fs');


function MixpanelRetention(apiKey, apiSecret) {

  this.apiKey = apiKey || env.MIXPANEL_API_KEY;
  this.apiSecret = apiSecret || env.MIXPANEL_API_SECRET;

  if (!this.apiKey) {
    throw new Error('API Key is required');
  }

  if (!this.apiSecret) {
    throw new Error('API Secret is required');
  }

  this.client = new MixpanelExport({
    api_key: apiKey,
    api_secret: apiSecret,
    timeout_after: 20
  });

}

MixpanelRetention.prototype.retentionReport = function(parameters) {

  var WORKSHEET1 = 'xl/worksheets/sheet1.xml';
  var zip;

  var parameters = parameters || {};

  var bornEvent = parameters.bornEvent;
  if (!bornEvent) {
    throw new Error('Born Event is required');
  }

  if (!parameters.startDate) {
    throw new Error('Start Date is required');
  }
  var startDate = Moment(parameters.startDate);

  var unit = parameters.unit || 'week';
  if (!(['day', 'week', 'month'].includes(unit))) {
    throw new Error('Unit must be day, week, or month');
  }

  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate()-1);
  var endDate = Moment(parameters.endDate || yesterday);
  if (startDate.isAfter(endDate)) {
    throw new Error('Start Date cannot be greater than End Date');
  }

  var duration = Moment.duration(endDate.diff(startDate));
  var intervals = 1;
  if (unit == 'day') {
    intervals = Math.ceil(duration.asDays());
  } else if (unit == 'week') {
    intervals = Math.ceil(duration.asWeeks());
  } else {
    intervals = Math.ceil(duration.asMonths());
  }
  var intervalCount = Math.min(intervals, 30);

  var style = parameters.style || {};
  var fontName        = style.fontName         || 'Helvetica';
  var headerFontSize  = style.headerFontSize   || '13';
  var headerFgColor   = style.headerFgColor    || '47596b';
  var headerBgColor   = style.headerBgColor    || 'f4f9fd';

  var retentionOptions = {
    from_date: startDate.format('YYYY-MM-DD'),
    to_date: endDate.format('YYYY-MM-DD'),
    retention_type: 'birth',
    born_event: bornEvent,
    unit: unit,
    interval_count: intervalCount
  };
  if (parameters.where) {
    retentionOptions.where = parameters.where;
  }
  if (parameters.bornWhere) {
    retentionOptions.born_where = parameters.bornWhere;
  }
  if (parameters.event) {
    retentionOptions.event = parameters.event;
  }

  return this.client.retention(retentionOptions)
    .then(function(data) {
      var unitAbrev = unit[0];

      var headers = [
        'Segment',
        'People',
        '< 1' + unitAbrev
      ].concat(_.map(_.range(1, intervalCount+1),
        function(i) {
          return i + unitAbrev
        }));
      var aoa = _(data)
        .mapValues(function(value, date) {
          return _.merge({}, value, {date});
        })
        .values()
        .sortBy('date')
        .map(function(m) {
          m.counts.pop();
          var percentages = _(m.counts)
            .map(function(c) {
              return (c / m.first).toFixed(2);
            })
            .filter(function(p) {
              return p >= 0.00;
            })
            .value();
          return [
            Moment(m.date).format('MMM D, YYYY'),
            m.first
          ].concat(percentages);
        })
        .value();
      var table = [headers].concat(aoa);

      var wb = { SheetNames:[], Sheets:{} };
      var ws = {};

      var range = {s: {c:0, r:0}, e: {c:0, r:0 }};

      for(var R = 0; R != table.length; ++R) {
        if(range.e.r < R) range.e.r = R;
        for(var C = 0; C != table[R].length; ++C) {
          if(range.e.c < C) range.e.c = C;

          var cell = { v: table[R][C] };
          if(cell.v == null) continue;

          var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

          cell.s = { alignment: { horizontal: 'center', vertical: 'center' }};
          if (R == 0 || C == 0) {
            cell.t = 's';
            cell.z = '@';
            cell.s.font = { sz: headerFontSize, color: { rgb: headerFgColor }};
            if (R == 0) {
              cell.s.font.bold = true;
            }
            if (C >= 2) {
              cell.s.fill = {
                fgColor: { rgb: 'FFFCE7' }
              };
            } else {
              cell.s.fill = {
                fgColor: { rgb: headerBgColor }
              };
            }
          } else if (C == 1) {
            cell.t = 'n';
            cell.z = '#,##0';
            cell.s.font = { sz: headerFontSize, color: { rgb: headerFgColor }};
            cell.s.fill = {
              fgColor: { rgb: headerBgColor }
            };
          } else {
            cell.s.font = { color: { rgb: 'ffffff'}};
            cell.t = 'n';
            cell.z = '0%';
          }
          cell.s.border = {};
          if (R == 0) {
            cell.s.border.bottom = { style: 'thin', color: { rgb: headerFgColor }};
          }
          if (C == 1) {
            cell.s.border.right = { style: 'thin', color: { rgb: headerFgColor }};
          }
          cell.s.font.name = fontName;

          ws[cell_ref] = cell;

          XLSX.utils.format_cell(ws[cell_ref]);
        }
      }
      ws['!ref'] = XLSX.utils.encode_range(range);
      ws['!cols'] = [{wch:18},{wch:12}].concat(
        _.times(intervalCount+2, function () {
          return {wch:5.5};
        })
      );
      var wsName = 'Retention';
      wb.SheetNames.push(wsName);
      wb.Sheets[wsName] = ws;

      var writeOptions = {
        showGridLines: false,
        bookType: 'xlsx',
        bookSST: false,
        type:'binary'
      };

      var wbout = XLSX.write(wb, writeOptions);
      return Jszip.loadAsync(wbout);
    })
    .then(function(ziper) {
      zip = ziper;
      return zip.file(WORKSHEET1).async("string");
    })
    .then(function (worksheet) {
      var cell_ref = XLSX.utils.encode_cell({c:intervalCount+2,r:intervals});
      var cond = '</sheetData><conditionalFormatting sqref="C2:' + cell_ref + '"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min" /><cfvo type="percentile" val="70" /><cfvo type="max" /><color rgb="FFC00000" /><color theme="3" tint="0.40" /><color theme="3" tint="-0.25" /></colorScale></cfRule></conditionalFormatting>';
      worksheet = worksheet.replace('</sheetData>', cond);
      zip.file(WORKSHEET1,worksheet);
      return zip.generateNodeStream({type:'nodebuffer',streamFiles:true});
    });

};

module.exports = MixpanelRetention;
