/* global window, document, Blob */
import React, { Component } from 'react';
import PropTypes from 'prop-types';

const propTypes = {
  table: PropTypes.string.isRequired,
  filename: PropTypes.string.isRequired,
  filetype: PropTypes.string,
  sheet: PropTypes.string.isRequired,
  id: PropTypes.string,
  className: PropTypes.string,
  buttonText: PropTypes.string,
};

const defaultProps = {
  id: 'button-download-as-xls',
  className: 'button-download',
  buttonText: 'Download',
};

class ReactHTMLTableToExcel extends Component {
  constructor(props) {
    super(props);
    this.handleDownload = this.handleDownload.bind(this);
  }

  static base64(s) {
    return window.btoa(unescape(encodeURIComponent(s)));
  }

  static format(s, c) {
    return s.replace(/{(\w+)}/g, (m, p) => c[p]);
  }

  handleDownload() {
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

    const table = document.getElementById(this.props.table).outerHTML;
    const sheet = String(this.props.sheet);
    const filetype = String((this.props.filetype) ? (this.props.filetype === 'xlsx' ? 'xlsx' : 'xls') : 'xls');
    const filename = `${String(this.props.filename) + '.' + filetype}`;

    const uri = 'data:application/vnd.ms-excel;base64,';
    const template =
      '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body>{table}</body></html>';

    const context = {
      worksheet: sheet || 'Worksheet',
      table,
    };

    // If IE11 or Edge
    if (window.navigator.msSaveOrOpenBlob) {
      const fileData = [
        `${'<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body>'}${table}</body></html>`,
      ];
      const blobObject = new Blob(fileData, { type: 'application/vnd.ms-excel' });
      window.navigator.msSaveOrOpenBlob(blobObject, filename);

      return true;
    }

    // For modern browsers
    const element = window.document.createElement('a');
    const content = (filetype === 'xlsx')
      ? ReactHTMLTableToExcel.format(template, context)
      : ReactHTMLTableToExcel.format(template, context);  // The logic remains mostly the same for styling in XLSX format (as styling is applied with HTML)

    element.href = uri + ReactHTMLTableToExcel.base64(content);
    element.download = filename;
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);

    return true;
  }

  render() {
    return (
      <button
        id={this.props.id}
        className={this.props.className}
        type="button"
        onClick={this.handleDownload}
      >
        {this.props.children ? this.props.children : this.props.buttonText}
      </button>
    );
  }
}

ReactHTMLTableToExcel.propTypes = propTypes;
ReactHTMLTableToExcel.defaultProps = defaultProps;

export default ReactHTMLTableToExcel;
