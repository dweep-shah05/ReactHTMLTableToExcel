'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _react = require('react');

var _react2 = _interopRequireDefault(_react);

var _propTypes = require('prop-types');

var _propTypes2 = _interopRequireDefault(_propTypes);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; } /* global window, document, Blob */


var propTypes = {
  table: _propTypes2.default.string.isRequired,
  filename: _propTypes2.default.string.isRequired,
  sheet: _propTypes2.default.string.isRequired,
  id: _propTypes2.default.string,
  className: _propTypes2.default.string,
  buttonText: _propTypes2.default.string
};

var defaultProps = {
  id: 'button-download-as-xlsx',
  className: 'button-download',
  buttonText: 'Download'
};

var ReactHTMLTableToExcel = function (_Component) {
  _inherits(ReactHTMLTableToExcel, _Component);

  function ReactHTMLTableToExcel(props) {
    _classCallCheck(this, ReactHTMLTableToExcel);

    var _this = _possibleConstructorReturn(this, (ReactHTMLTableToExcel.__proto__ || Object.getPrototypeOf(ReactHTMLTableToExcel)).call(this, props));

    _this.handleDownload = _this.handleDownload.bind(_this);
    return _this;
  }

  // Helper function to convert HTML table to CSV


  _createClass(ReactHTMLTableToExcel, [{
    key: 'tableToCSV',
    value: function tableToCSV() {
      var table = document.getElementById(this.props.table);
      var rows = table.querySelectorAll('tr');
      var csv = '';

      rows.forEach(function (row) {
        var cols = row.querySelectorAll('td, th');
        var rowData = [];
        cols.forEach(function (col) {
          var cellText = col.innerText.trim();
          cellText = cellText.replace(/"/g, '""'); // Escape any quotes
          if (cellText.includes(',')) {
            cellText = '"' + cellText + '"'; // Wrap in quotes if it contains commas
          }
          rowData.push(cellText);
        });
        csv += rowData.join(',') + '\r\n'; // Join with commas, add new line for each row
      });

      return csv;
    }

    // Handle file download logic

  }, {
    key: 'handleDownload',
    value: function handleDownload() {
      if (!document) {
        if (process.env.NODE_ENV !== 'production') {
          console.error('Failed to access document object');
        }
        return null;
      }

      if (document.getElementById(this.props.table).nodeType !== 1 || document.getElementById(this.props.table).nodeName !== 'TABLE') {
        if (process.env.NODE_ENV !== 'production') {
          console.error('Provided table property is not an HTML table element');
        }
        return null;
      }

      // Generate CSV data from the HTML table
      var csv = this.tableToCSV();

      // Mimic an Excel file using CSV data in a .xlsx wrapper
      var uri = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,';
      var template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' + '<head><meta charset="UTF-8"><!--[if gte mso 9]>' + '<xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name>' + '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>' + '</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head>' + '<body><pre>{csv}</pre></body></html>';

      var context = {
        worksheet: this.props.sheet || 'Sheet1',
        csv: csv // Insert the CSV data as body content
      };

      var filename = String(this.props.filename) + '.xlsx';

      // For Internet Explorer (IE11 support)
      if (window.navigator.msSaveOrOpenBlob) {
        var fileData = ['' + ('<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><!--[if gte mso 9]>' + '<xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name>' + '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>' + '</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><pre>') + csv + '</pre></body></html>'];
        var blobObject = new Blob(fileData, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        window.navigator.msSaveOrOpenBlob(blobObject, filename);
        return true;
      }

      // For modern browsers
      var element = document.createElement('a');
      element.href = uri + ReactHTMLTableToExcel.base64(ReactHTMLTableToExcel.format(template, context));
      element.download = filename;
      document.body.appendChild(element);
      element.click();
      document.body.removeChild(element);

      return true;
    }
  }, {
    key: 'render',
    value: function render() {
      return _react2.default.createElement(
        'button',
        {
          id: this.props.id,
          className: this.props.className,
          type: 'button',
          onClick: this.handleDownload
        },
        this.props.buttonText
      );
    }
  }]);

  return ReactHTMLTableToExcel;
}(_react.Component);

ReactHTMLTableToExcel.propTypes = propTypes;
ReactHTMLTableToExcel.defaultProps = defaultProps;

exports.default = ReactHTMLTableToExcel;
