const { GoogleSpreadsheet } = require('google-spreadsheet');
const credentials = require('./credentials.json');
const arquivo = require('./doc.json');

//Create a new Spreadsheet doc
const getDoc = async () => {
    const doc = new GoogleSpreadsheet(arquivo.id);

    await doc.useServiceAccountAuth({
        client_email: credentials.client_email,
        private_key: credentials.private_key.replace(/\\n/g, '\n')
    })
    await doc.loadInfo();
    return doc;
}

//Show doc title
getDoc().then(doc => {
    console.log("Document loaded " + doc.title);
});

//Function to round to next integer
function format(n) {
    if (n % 1 == 0) {
        return n + "";
    }

    else return n - (n % 1) + 1 + "";
}

let sheet;
getDoc().then(doc => {
    sheet = doc.sheetsByIndex[0]; //Choose the sheet to be manipulated
    sheet.loadHeaderRow(3); //Determines the header Row 

    //Get all rows in the worksheet
    sheet.getRows().then(rows => {

        //For every row determines the student situation
        rows.map(row => {
            console.log("Calculating " + row.Aluno + "'s average");
            let m = (row.P1 * 1 + row.P2 * 1 + row.P3 * 1) / 30; //Average calculation
            console.log("Determining " +row.Aluno + "'s situation");
            if (row.Faltas > 0.25 * 60) {
                row.Situação = "Reprovado por Falta";
                row.Nota_para_Aprovação_Final = format(0);

            }
            else {
                if (m < 5) {
                    row.Situação = "Reprovado por Nota";
                    row.Nota_para_Aprovação_Final = "0";
                }
                else {
                    if (m < 7 && 5 <= m) {
                        row.Situação = "Exame Final"
                        let naf = (10 - m); //"Nota final" calculation
                        row.Nota_para_Aprovação_Final = format(naf);
                    }
                    else {
                        row.Situação = "Aprovado";
                        row.Nota_para_Aprovação_Final = "0";
                    }
                }
            }

            //Save changes on worksheet
            row.save().then(() => {
                console.log(row.Aluno + " student information updated");
            }

            );
        });
    })
});


