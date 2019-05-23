
var Promise = XlsxPopulate.Promise;

function getWorkbook() {
    return new Promise(function (resolve, reject) {
        var req = new XMLHttpRequest();
        var url = './gp.xlsx';
        req.open("GET", url, true);
        req.responseType = "arraybuffer";
        req.onreadystatechange = function () {
            if (req.readyState === 4) {
                if (req.status === 200) {
                    resolve(XlsxPopulate.fromDataAsync(req.response))
                } else {
                    reject('Received a ' + req.status + ' HTTP code.')
                }
            }
        }

        req.send();
    })
}

function generate(type) {
    return getWorkbook()
        .then(function (workbook) {

            // Get inputs
            const name = document.getElementById('name');
            const allRiders = [...document.querySelectorAll('.gp__rider > .gp__rider__input')];

            let ridersCellStart = 4;

            // Set hosts name cell value
            workbook.sheet(0).cell('B2').value(name.value);

            // Loop through all riders cells and set value
            // for (let i = 0; i <= 17; i++) {
            //     workbook.sheet(0).cell(`H${ridersCellStart + 2}`).value(allRiders[i].value);
            // }

            workbook.sheet(0).cell('H4').value(allRiders[0].value);
            workbook.sheet(0).cell('H6').value(allRiders[1].value);
            workbook.sheet(0).cell('H8').value(allRiders[2].value);
            workbook.sheet(0).cell('H10').value(allRiders[3].value);
            workbook.sheet(0).cell('H12').value(allRiders[4].value);
            workbook.sheet(0).cell('H14').value(allRiders[5].value);
            workbook.sheet(0).cell('H16').value(allRiders[6].value);
            workbook.sheet(0).cell('H18').value(allRiders[7].value);

            workbook.sheet(0).cell('H20').value(allRiders[8].value);
            workbook.sheet(0).cell('H22').value(allRiders[9].value);
            workbook.sheet(0).cell('H24').value(allRiders[10].value);
            workbook.sheet(0).cell('H26').value(allRiders[11].value);
            workbook.sheet(0).cell('H28').value(allRiders[12].value);
            workbook.sheet(0).cell('H30').value(allRiders[13].value);
            workbook.sheet(0).cell('H32').value(allRiders[14].value);
            workbook.sheet(0).cell('H34').value(allRiders[15].value);
            workbook.sheet(0).cell('H36').value(allRiders[15].value);
            workbook.sheet(0).cell('H38').value(allRiders[15].value);

            return workbook.outputAsync({ type: type })
        });
}

function generateBlob(e) {
    e.preventDefault();
    return generate()
        .then(function (blob) {
            var url = window.URL.createObjectURL(blob);
            var a = document.createElement('a');
            document.body.appendChild(a);
            a.href = url;
            a.download = 'Program Żużlowy.xlsx';
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        })
}

document.querySelector('.generator__form').addEventListener('submit', generateBlob);
