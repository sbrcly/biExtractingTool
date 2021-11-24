const todaysGames = document.querySelector('.todays-games');
const gamesInput = document.querySelector('textarea');
const teamsList = document.querySelector('.teams-list');
const teamsList2 = document.querySelector('.teams-list2');
const getBis = document.querySelector('#getBis');

const dbLink = document.querySelector('#dbLink');
dbLink.setAttribute('href', 'https://www.donbest.com/schedules');
dbLink.setAttribute('target', '_blank');

let dbGamesFinal = [];
let excelBtn;

const nameVarients = [
    {
        espn: 'Green Bay',
        db: 'WISC GREEN BAY'
    },
    {
        espn: 'UMass',
        db: 'MASSACHUSETTS'
    },
    {
        espn: 'Sam Houston',
        db: 'SAM HOUSTON ST'
    },
    {
        espn: 'Texas A&M-CC',
        db: 'TEXAS A&M CORPUS'
    },
    {
        espn: 'UNC Greensboro',
        db: 'NC GREENSBORO'
    },
    {
        espn: 'Louisiana',
        db: 'UL - LAFAYETTE'
    },
    {
        espn: 'Northern Colorado',
        db: 'NO. COLORADO'
    }
]

getBis.addEventListener('click', () => {
    const donBestGames = gamesInput.value;
    const parseGames = donBestGames.split('\n');
    const dbGames = parseGames.filter(game => {
    let item = game.split(' ');
    if (item.length === 1) {
        return item[0].length > 2 &&
                item[0].length <= 6 &&
                parseInt(item[0]) % 1 === 0;
    }   else {
        return item[0].length <= 6 && parseInt(item[0]) % 1 === 0;
    }
    
    });
    
    for (let game of dbGames) {
        const splitGame = game.split(' ');
        if (!splitGame[1]) {
            game = game + ' TBD';
            dbGamesFinal.push(game);
        }   else {
            dbGamesFinal.push(game);
        }
    }
    
    const oddBis = dbGamesFinal.filter(game => {
        const bi = game.split(' ')[0];
        return Number(bi) % 2 !== 0;
    });

    const evenBis = dbGamesFinal.filter(game => {
        const bi = game.split(' ')[0];
        return Number(bi) % 2 === 0;
    });

    const tables = document.querySelectorAll('table');

    for (let table of tables) {
        const tableHeading = document.createElement('tr');
        tableHeading.classList.add('tableHeading');
    
        const awayBIs = document.createElement('td');
        awayBIs.innerText = 'Away BI';
    
        const awayTeams = document.createElement('td');
        awayTeams.innerText = 'Away Team';
    
        const homeBIs = document.createElement('td');
        homeBIs.innerText = 'Home BI';
    
        const homeTeams = document.createElement('td');
        homeTeams.innerText = 'Home Team';
    
        tableHeading.append(awayBIs, awayTeams, homeBIs, homeTeams);
        table.append(tableHeading);
        table.classList.add('show-table');
    };

    for (let i = 0; i < oddBis.length; i++) {
        const newGame = document.createElement('tr');

        const awayBI = document.createElement('td');
        awayBI.classList.add('awayBI');
        awayBI.innerText = oddBis[i].split(' ')[0];

        const awayTeam = document.createElement('td');
        awayTeam.classList.add('awayTeam');
        awayTeam.innerText = oddBis[i].split(' ').slice(1, oddBis[i].length).join(' ');

        const homeBI = document.createElement('td');
        homeBI.classList.add('homeBI');
        homeBI.innerText = evenBis[i].split(' ')[0];

        const homeTeam = document.createElement('td');
        homeTeam.classList.add('homeTeam');
        homeTeam.innerText = evenBis[i].split(' ').slice(1, evenBis[i].length).join(' ');

        newGame.append(awayBI, awayTeam, homeBI, homeTeam);
        if (i < oddBis.length / 2) {
            teamsList.append(newGame);
        }   else {
            teamsList2.append(newGame);
        }
        
    }
    const sendToExcel = document.createElement('button');
    sendToExcel.classList.add('sendToExcel');
    sendToExcel.innerText = 'Export to Excel';

    todaysGames.append(sendToExcel);
    getBis.disabled = 'true';
    excelBtn = document.querySelector('.sendToExcel');

});



// EXCEL SHEET MODIFICATION

let wbData = []

function handleFile(e) {
    var files = e.target.files, f = files[0];
    var reader = new FileReader();
    
    reader.onload = function(e) {
        var workbook = XLSX.read(e.target.result);

        /* GET S COLUMN VALUES FROM CBB ESPN WORKSHEET */
        const worksheet = workbook.Sheets['CBB ESPN'];
        for (let i = 0; i < 100; i++) {
            if (workbook.Sheets['CBB ESPN'][`S${i}`]) {
                wbData.push(worksheet[`S${i}`].v);
            };
        };
        // CREATE AN ARRAY OF OBJECTS WITH THE DB BIS MATCHED WITH THE TEAMS IN THE S CELL
        let finalBis = [];
        excelBtn.addEventListener('click', () => {
            for (let i = 0; i < wbData.length; i++) {
                for (let game of dbGamesFinal) {
                    const splitGame = game.split(' ');
                    let parseTeam = [];
                    for (let i = splitGame.length - 1; i > 0; i--) {
                        parseTeam.unshift(splitGame[i]);
                    }
                    const teamName = parseTeam.join(' ');
                    if (teamName === (wbData[i].toUpperCase()) &&
                        wbData[i].toUpperCase() != 'TBD' ) {
                        finalBis.push(
                            {
                                bi: game.split(' ')[0],
                                game: wbData[i]
                            })
                    }   else {
                        for (let name of nameVarients) {
                            if (teamName === name.db) {
                                if (wbData[i] === name.espn) {
                                    finalBis.push(
                                        {
                                            bi: game.split(' ')[0],
                                            game: wbData[i]
                                        })
                                }
                            }
                        }
                    }
                }
            }
            // PUT THE DB BIS INTO THE Z COLUMN
            for (let i = 0; i < 100; i++) {
                if (workbook.Sheets['CBB ESPN'][`S${i}`]) {
                    for (let final of finalBis) {
                        if (workbook.Sheets['CBB ESPN'][`S${i}`].v === final.game) {
                            workbook.Sheets['CBB ESPN'][`Z${i}`] = {
                                t: 'n',
                                v: final.bi,
                                w: final.bi
                            }
                        }
                    }
                };
            };
            console.log(worksheet);
            XLSX.writeFile(workbook, 'auto-schedule.xlsx');
            const successMsg = document.createElement('p');
            successMsg.innerText = 'Export successful'
            setTimeout(() => {
                todaysGames.append(successMsg);
            }, 1000);
        });
            
    };
    reader.readAsArrayBuffer(f);
};
    
const input = document.querySelector('#input');
input.addEventListener('change', handleFile, false);