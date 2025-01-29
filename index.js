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

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var propTypes = {
  table: _propTypes2.default.string.isRequired,
  filename: _propTypes2.default.string.isRequired,
  id: _propTypes2.default.string,
  className: _propTypes2.default.string,
  buttonText: _propTypes2.default.string
};

var defaultProps = {
  id: 'button-download-as-csv',
  className: 'button-download',
  buttonText: 'Download CSV'
};

var ReactHTMLTableToCSV = function (_Component) {
  _inherits(ReactHTMLTableToCSV, _Component);

  function ReactHTMLTableToCSV(props) {
    _classCallCheck(this, ReactHTMLTableToCSV);

    var _this = _possibleConstructorReturn(this, (ReactHTMLTableToCSV.__proto__ || Object.getPrototypeOf(ReactHTMLTableToCSV)).call(this, props));

    _this.handleDownload = _this.handleDownload.bind(_this);
    return _this;
  }

  _createClass(ReactHTMLTableToCSV, [{
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
          console.error('Provided table property is not html table element');
        }
        return null;
      }

      var table = document.getElementById(this.props.table);
      var rows = table.querySelectorAll('tr');
      var csv = '';

      rows.forEach(function (row) {
        var cols = row.querySelectorAll('td, th');
        var rowData = [];
        cols.forEach(function (col) {
          // Escape any commas and quotes in the cell data
          var cellText = col.innerText.trim();
          cellText = cellText.replace(/"/g, '""'); // Double quotes inside quotes
          if (cellText.includes(',')) {
            cellText = `"${cellText}"`; // Wrap in quotes if it contains commas
          }
          rowData.push(cellText);
        });
        csv += rowData.join(',') + '\r\n'; // Join cell data with commas and add a new line
      });

      // Create a link element and trigger the download
      var filename = this.props.filename + '.csv';
      var encodedUri = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csv);
      var element = document.createElement('a');
      element.setAttribute('href', encodedUri);
      element.setAttribute('download', filename);
      document.body.appendChild(element);
      element.click();
      document.body.removeChild(element);
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

  return ReactHTMLTableToCSV;
}(_react.Component);

ReactHTMLTableToCSV.propTypes = propTypes;
ReactHTMLTableToCSV.defaultProps = defaultProps;

exports.default = ReactHTMLTableToCSV;
