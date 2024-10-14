import React from 'react';
import './App.css';
import ExcelImporter from './components/ExcelImporter';
import{ Typography }from"antd";

const { Title } = Typography;
function App() {
  
  return (
    <div className="App">
      <Title>Importador de Excel</Title>
      <ExcelImporter />
    </div>
  );
}

export default App;
