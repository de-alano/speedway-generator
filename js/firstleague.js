
var Promise = XlsxPopulate.Promise;

function getWorkbook() {
    return new Promise(function (resolve, reject) {
        var req = new XMLHttpRequest();
        var url = './1liga.xlsx';
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
            const hosts = document.getElementById('hosts');
            const guests = document.getElementById('guests');
            const allRiders = [...document.querySelectorAll('.team__rider > .team__rider__input')];

            let guestsCellStart = 5;
            let hostsCellStart = 5;

            // Set hosts name cell value
            workbook.sheet(0).cell('C2').value(hosts.value);

            // Loop through all hosts cells and set value
            for (let i = 0; i <= 7; i++) {
                workbook.sheet(0).cell(`C${guestsCellStart++}`).value(allRiders[i].value);
            }

            // Set quests name cell value
            workbook.sheet(0).cell('U2').value(guests.value);

            // Loop through all guests cells and set value
            for (let i = 8; i <= 15; i++) {
                workbook.sheet(0).cell(`U${hostsCellStart++}`).value(allRiders[i].value);
            }

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