const XLSX = require('xlsx');

const ExcelAJSON = () => {
    const excel = XLSX.readFile('ReporteCV6011.xls');
    const nombreHoja = excel.SheetNames; // regresa un array
    let datos = XLSX.utils.sheet_to_json(excel.Sheets[nombreHoja[0]]);
    
    let json_data_temp = {};
    let json_data_new = [];

    for (let i = 0; i < datos.length; i++) {
        const dato = datos[i];
        // Reemplazar espacios por guion bajo en el nombre del campo y pasar todo a minúsculas
        const nombreCampo = dato['Tipo'].split(' ').join('_').toLowerCase();
        const datoCampo = dato['Dato'];
        if(nombreCampo === 'bank_reference') {
            // CREAR OBJETO NUEVO
            if( json_data_temp.length === 0 ) {
                // Detección de primer Bank Reference
                json_data_temp = { ...json_data_temp, [nombreCampo] : datoCampo };
            } else {
                json_data_new.push(json_data_temp);
                json_data_temp = {};
                json_data_temp = { ...json_data_temp, [nombreCampo] : datoCampo };
            }
        } else {
            json_data_temp = { ...json_data_temp, [nombreCampo] : datoCampo };
        }

    }

    console.log(json_data_new);
}

ExcelAJSON();