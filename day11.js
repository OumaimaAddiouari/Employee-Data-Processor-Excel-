//impote le module xlsx qui est utiliser pouer lire et ecrire des fichier excel dans node 
const XLSX = require('xlsx');
//pour le remplacement 
const path = require('path');
// for read file 
const file = './employee_data_.xlsx';
readFile(file);

try {
    function readFile(filePath) {
        //lit le fichier Excel spécifie par "filepath" et stocké son contenu dans 'workbook'
        const workbook = XLSX.readFile(filePath);
        //Obtient le nom de la première feuille dans le classeur.
        const sheetName = workbook.SheetNames[0];
        // Obtient la feuille de calcul en utilisant le nom de la feuille.    
        const worksheet = workbook.Sheets[sheetName];
        // Convertit la feuille de calcul en format JSON.
        const data = XLSX.utils.sheet_to_json(worksheet);
        console.log(data);
    }





    function CalculatingBonuses(empleyee) {
        if (empleyee.AnnualSalary < 50000) {
            return empleyee.BonusePercentage = 5;
        } else if (empleyee.AnnualSalary >= 50000 && salary <= 100000) {
            return empleyee.BonusePercentage = 7;
        } else {
            return empleyee.BonusePercentage = 10;
        }
        empleyee.BonusAmount = (empleyee.BonusePercentage / 100) * empleyee.AnnualSalary;
        return empleyee.BonusAmount;
    }

    //create a new worksheet(feuille) 

    const newSheet = XLSX.utils.json_to_sheet(data);


    // Add  two new columns

    XLSX.utils.sheet_add_aoa(newSheet, [['BonusePercentage'], ['BonusAmount']])

    // write the results to a new Excel file.
    //create a new workbook  with the new sheet

    const newWorkbook = xlsx.utils.book_new();

    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "shhet1");

    const newFilePath = path.join(__dirname, 'employee_data_with_bonus.xlsx');

    XLSX.writeFile(newWorkbook, newFilePath);
    console.log('New excel file : ${newFilePath}');


} catch (error) {
    console.error("error:" + error.message);
}



