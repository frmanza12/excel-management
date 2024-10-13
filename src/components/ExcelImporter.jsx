import React, { useState } from 'react';
import { Upload, Button, Table, DatePicker, Select, Spin, Row, Col } from 'antd';
import { UploadOutlined } from '@ant-design/icons';
import * as XLSX from 'xlsx';
import moment from 'moment';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
const { RangePicker } = DatePicker;
const { Option } = Select;

const convertExcelDate = (excelDate) => {
  if (typeof excelDate === 'number') {
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    if (date.getFullYear() >= 1900 && date.getFullYear() <= 2100) {
      return date;
    }
  }
  return excelDate;
};

const ExcelImporter = () => {
  const [data, setData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [dateFilters, setDateFilters] = useState({});
  const [appliedDateFilters, setAppliedDateFilters] = useState({});
  const [technicians, setTechnicians] = useState([]);
  const [selectedTechnicians, setSelectedTechnicians] = useState([]);
  const [reportStatus, setReportStatus] = useState(null);
  const [loading, setLoading] = useState(false);
  const [filtering, setFiltering] = useState(false);
// Función para dividir un array en fragmentos de un tamaño especificado
const chunkArray = (array, size) => {
  const result = [];
  for (let i = 0; i < array.length; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
};

  const handleFileUpload = (file) => {
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      const bstr = e.target.result;
      const workbook = XLSX.read(bstr, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
      const dateColumns = detectDateColumns(jsonData).filter((colName) => colName !== 'Id');

      let allRows = [];
      const chunks = chunkArray(jsonData.slice(1), 2000);
      chunks.forEach((chunk, chunkIndex) => {
        const rows = chunk.map((row, index) => {
          const rowData = {};
          jsonData[0].forEach((col, i) => {
            const cellValue = row[i];
            rowData[col] = dateColumns.includes(i) && typeof cellValue === 'number'
              ? convertExcelDate(cellValue) || cellValue
              : cellValue;
          });
          return { key: chunkIndex * 2000 + index, ...rowData };
        });
        allRows = allRows.concat(rows);
      });

      const cols = jsonData[0].map((col, index) => ({
        title: col,
        dataIndex: col,
        key: col,
        isDate: dateColumns.includes(index),
        render: (text) => {
          if (text instanceof Date) {
            return moment(text).format('DD/MM/YYYY');
          }
          return text || "";
        },
      }));

      const technicianColumnIndex = jsonData[0].indexOf("Técnico");
      if (technicianColumnIndex !== -1) {
        const technicianList = [...new Set(allRows.map(row => row["Técnico"]).filter(Boolean))];
        technicianList.sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' })); // Orden alfabético
        setTechnicians(technicianList);
      }

      setColumns(cols);
      setData(allRows);
      setDateFilters({});
      setAppliedDateFilters({});
      setSelectedTechnicians([]);
      setReportStatus(null);
      setLoading(false);
    };
    reader.readAsBinaryString(file);
    return false;
  };

  const detectDateColumns = (jsonData) => {
    const dateColumns = [];
    jsonData[0].forEach((col, index) => {
      const sampleValues = jsonData.slice(1, 11).map(row => row[index]);
      if (col === 'Id') return;
      const validDatesCount = sampleValues.filter(value =>
        value !== undefined && value !== null && typeof value === 'number' && convertExcelDate(value)
      ).length;
      if (validDatesCount / sampleValues.filter(value => value !== undefined && value !== null).length >= 0.8) {
        dateColumns.push(index);
      }
    });

    return dateColumns;
  };

  // Maneja el cambio en el filtro de rango de fechas
  const handleDateRangeFilterChange = (columnName, dates) => {
    setDateFilters((prevFilters) => ({
      ...prevFilters,
      [columnName]: dates ? [dates[0].startOf('day'), dates[1].endOf('day')] : null,
    }));
  };

  const applyDateFilters = () => {
    setFiltering(true);
    setAppliedDateFilters(dateFilters);
    setTimeout(() => setFiltering(false), 500);
  };

  const exportFilteredDataWithClassification = async () => {
    const filteredData = getFilteredData();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Estadísticas');
  
    // Tabla: Clasificación de Entradas y Salidas por Año y Mes
    worksheet.addRow(['Clasificación de Entradas y Salidas por Año y Mes']);
    worksheet.addRow(['Año', 'Mes', 'Entradas', 'Salidas']);
  
    const classificationData = {};
  
    filteredData.forEach((row) => {
      const entryDate = row['Fecha Petición'] instanceof Date ? row['Fecha Petición'] : null;
      const exitDate = row['Fecha Informe'] instanceof Date ? row['Fecha Informe'] : null;
  
      if (entryDate) {
        const year = entryDate.getFullYear();
        const month = entryDate.getMonth() + 1;
        if (!classificationData[year]) classificationData[year] = {};
        if (!classificationData[year][month]) classificationData[year][month] = { entradas: 0, salidas: 0 };
        classificationData[year][month].entradas += 1;
      }
  
      if (exitDate) {
        const year = exitDate.getFullYear();
        const month = exitDate.getMonth() + 1;
        if (!classificationData[year]) classificationData[year] = {};
        if (!classificationData[year][month]) classificationData[year][month] = { entradas: 0, salidas: 0 };
        classificationData[year][month].salidas += 1;
      }
    });
  
    const monthNames = [
      '', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
      'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
    ];
  
    Object.entries(classificationData).forEach(([year, months]) => {
      let firstRow = true;
      let yearEntradasTotal = 0;
      let yearSalidasTotal = 0;
  
      Object.entries(months).forEach(([month, counts]) => {
        worksheet.addRow([
          firstRow ? year : '', // Solo muestra el año en la primera fila
          monthNames[parseInt(month)], // Convertir número de mes a nombre
          counts.entradas,
          counts.salidas,
        ]);
  
        yearEntradasTotal += counts.entradas;
        yearSalidasTotal += counts.salidas;
        firstRow = false;
      });
  
      worksheet.addRow([
        '', 
        'Total', 
        yearEntradasTotal, 
        yearSalidasTotal,
      ]);
    });
  
    // Tabla adicional: Clasificación de registros pendientes por técnico
    worksheet.addRow([]);
    worksheet.addRow(['Clasificación de Registros Pendientes por Técnico']);
    worksheet.addRow(['Técnico', 'Registros Pendientes']);
  
    const pendingRecordsByTechnician = {};
  
    filteredData.forEach((row) => {
      const technician = row["Técnico"];
      const isPending = !row["Fecha Informe"];
  
      if (isPending && technician) {
        if (!pendingRecordsByTechnician[technician]) {
          pendingRecordsByTechnician[technician] = 0;
        }
        pendingRecordsByTechnician[technician] += 1;
      }
    });
  
    Object.entries(pendingRecordsByTechnician).forEach(([technician, count]) => {
      worksheet.addRow([technician, count]);
    });
  
    // Tabla adicional: Clasificación de registros sacados por técnico
    worksheet.addRow([]);
    worksheet.addRow(['Clasificación de Registros Sacados por Técnico']);
    worksheet.addRow(['Técnico', 'Registros Sacados']);
  
    const completedRecordsByTechnician = {};
  
    filteredData.forEach((row) => {
      const technician = row["Técnico"];
      const isCompleted = row["Fecha Informe"] instanceof Date;
  
      if (isCompleted && technician) {
        if (!completedRecordsByTechnician[technician]) {
          completedRecordsByTechnician[technician] = 0;
        }
        completedRecordsByTechnician[technician] += 1;
      }
    });
  
    Object.entries(completedRecordsByTechnician).forEach(([technician, count]) => {
      worksheet.addRow([technician, count]);
    });
  
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'Estadísticas.xlsx');
  };
  const getFilteredData = () => {
    return data.filter((row) => {
      const passesDateFilters = columns.every((col) => {
        if (col.isDate && appliedDateFilters[col.dataIndex]) {
          const rowDate = row[col.dataIndex];
          if (!(rowDate instanceof Date)) return false;

          const [startDate, endDate] = appliedDateFilters[col.dataIndex];
          return (!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate);
        }
        return true;
      });

      const passesTechnicianFilter = selectedTechnicians.length === 0 || selectedTechnicians.includes(row["Técnico"]);
      const passesReportStatusFilter =
        !reportStatus || reportStatus === 'Todos' ||
        (reportStatus === 'Pendiente' && !row["Fecha Informe"]) ||
        (reportStatus === 'Sacado' && row["Fecha Informe"]);

      return passesDateFilters && passesTechnicianFilter && passesReportStatusFilter;
    });
  };

  const filteredData = getFilteredData();

  return (
    <div style={{ padding: 20 }}>
      <Upload beforeUpload={handleFileUpload} showUploadList={false}>
        <Button icon={<UploadOutlined />}>Importar Excel</Button>
      </Upload>

      {loading ? (
        <Spin tip="Cargando datos..." style={{ margin: '20px 0' }} />
      ) : (
        <>
          <Row gutter={16} style={{ marginTop: 20, marginBottom: 20, display: 'flex', justifyContent: 'left', alignItems: 'center' }}>
            {columns.map((col) =>
              col.isDate ? (
                <Col key={col.dataIndex} style={{ marginBottom: 10, display: 'flex', flexDirection: 'column' }}>
                  <span>{col.title}</span>
                  <RangePicker
                    size="small"
                    placeholder={["Fecha Inicio", "Fecha Fin"]}
                    onChange={(dates) => handleDateRangeFilterChange(col.dataIndex, dates)}
                    format="DD/MM/YYYY"
                  />
                </Col>
              ) : null
            )}
          </Row>
          <Row style={{ marginBottom: 20 }} gutter={12}>
            {technicians.length > 0 && (
              <Col span={12} style={{ marginBottom: 10, display: 'flex', flexDirection: 'column' }}>
                <span>Filtro por Técnico</span>
                <Select
                  mode="multiple"
                  placeholder="Selecciona Técnico"
                  style={{ width: '100%' }}
                  onChange={setSelectedTechnicians}
                  value={selectedTechnicians}
                >
                  {technicians.map((tech) => (
                    <Option key={tech} value={tech}>
                      {tech}
                    </Option>
                  ))}
                </Select>
              </Col>
            )}
            <Col span={12} style={{ marginBottom: 10, display: 'flex', flexDirection: 'column' }}>
              <span>Estado del Informe</span>
              <Select
                placeholder="Selecciona Estado"
                style={{ width: '100%' }}
                onChange={setReportStatus}
                value={reportStatus}
              >
                <Option value="Todos">Todos</Option>
                <Option value="Pendiente">Pendiente</Option>
                <Option value="Sacado">Sacado</Option>
              </Select>
            </Col>
          </Row>
          <Row>
            <Col>
              <Button type="primary" onClick={applyDateFilters} style={{ marginTop: 10 }}>
                Aplicar Filtros
              </Button>
            </Col>
            <Col>
              <Button type="primary" onClick={exportFilteredDataWithClassification} style={{ marginTop: 10, marginLeft: 10 }}>
                Exportar Datos y Clasificación
              </Button>
            </Col>
          </Row>
          <div style={{ marginBottom: 10 }}>
            <strong>Total de registros: {filteredData.length}</strong>
          </div>
          {filtering ? (
            <Spin tip="Aplicando filtros..." style={{ margin: '20px 0' }} />
          ) : (
            <div style={{ overflowX: 'auto', maxWidth: '100%' }}>
              <Table columns={columns} dataSource={filteredData} pagination={false} scroll={{ x: 'max-content' }} />
            </div>
          )}
        </>
      )}
    </div>
  );
};

export default ExcelImporter;