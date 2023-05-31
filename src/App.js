import React from 'react';
import * as XLSX from 'xlsx';

class ExcelReadWrite extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      data: null,
    };
  }

  handleFileChange = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      this.setState({ data: jsonData });
    };
    reader.readAsArrayBuffer(file);
  };

  exportToExcel = () => {
    const { data } = this.state;
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const excelBuffer = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });
    const dataBlob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    const fileName = 'export.xlsx';
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {
      window.navigator.msSaveOrOpenBlob(dataBlob, fileName);
    } else {
      const url = window.URL.createObjectURL(dataBlob);
      const link = document.createElement('a');
      link.href = url;
      link.download = fileName;
      link.click();
      window.URL.revokeObjectURL(url);
    }
  };

  render() {
    const { data } = this.state;

    return (
      <div>
        <input type="file" onChange={this.handleFileChange} />
        <button onClick={this.exportToExcel} disabled={!data}>
          Export to Excel
        </button>
        {data && (
          <table>
            <thead>
              <tr>
                {data[0].map((header, index) => (
                  <th key={index}>{header}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.slice(1).map((row, index) => (
                <tr key={index}>
                  {row.map((cell, index) => (
                    <td key={index}>{cell}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    );
  }
}

export default ExcelReadWrite;
