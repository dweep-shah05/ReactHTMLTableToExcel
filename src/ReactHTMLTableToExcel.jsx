/* global window, document, Blob */
import React, { Component } from 'react';
import PropTypes from 'prop-types';

const propTypes = {
  table: PropTypes.string.isRequired,
  filename: PropTypes.string.isRequired,
  sheet: PropTypes.string.isRequired,
  id: PropTypes.string,
  className: PropTypes.string,
  buttonText: PropTypes.string,
};

const defaultProps = {
  id: 'button-download-as-xlsx',
  className: 'button-download',
  buttonText: 'Download',
};

class ReactHTMLTableToExcel extends Component {
  constructor(props) {
    super(props);
    this.handleDownload = this.handleDownload.bind(this);
  }

  // Helper function to convert HTML table to CSV
  tableToCSV() {
    const table = document.getElementById(this.props.table);
    const rows = table.querySelectorAll('tr');
    let csv = '';
    
    rows.forEach((row) => {
      const cols = row.querySelectorAll('td, th');
      const rowData = [];
      cols.forEach((col) => {
        let cellText = col.innerText.trim();
        cellText = cellText.replace(/"/g, '""'); // Escape any quotes
        if (cellText.includes(',')) {
          cellText = `"${cellText}"`; // Wrap in quotes if it contains commas
        }
        rowData.push(cellText);
      });
      csv += rowData.join(',') + '\r\n'; // Join with commas, add new line for each row
    });

    return csv;
  }

  // Handle file download logic
  handleDownload() {
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
    const csv = this.tableToCSV();

    // Mimic an Excel file using CSV data in a .xlsx wrapper
    const uri = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,';
    const template =
      '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta charset="UTF-8"><!--[if gte mso 9]>' +
      '<xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name>' +
      '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>' +
      '</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head>' +
      '<body><pre>{csv}</pre></body></html>';

    const context = {
      worksheet: this.props.sheet || 'Sheet1',
      csv, // Insert the CSV data as body content
    };

    const filename = `${String(this.props.filename)}.xlsx`;

    // For Internet Explorer (IE11 support)
    if (window.navigator.msSaveOrOpenBlob) {
      const fileData = [
        `${'<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><!--[if gte mso 9]>' +
        '<xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name>' +
        '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>' +
        '</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><pre>'}${csv}</pre></body></html>`
      ];
      const blobObject = new Blob(fileData, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      window.navigator.msSaveOrOpenBlob(blobObject, filename);
      return true;
    }

    // For modern browsers
    const element = document.createElement('a');
    element.href = uri + ReactHTMLTableToExcel.base64(ReactHTMLTableToExcel.format(template, context));
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
        {this.props.buttonText}
      </button>
    );
  }
}

ReactHTMLTableToExcel.propTypes = propTypes;
ReactHTMLTableToExcel.defaultProps = defaultProps;

export default ReactHTMLTableToExcel;
