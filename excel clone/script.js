const PS = new PerfectScrollbar("#cells", {
    wheelSpeed: 12,
    wheelPropagation: true,
});

function findRowCOl(ele) {
    let idArray = $(ele).attr("id").split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);
    return [rowId, colId];
}

function calcColName(n){
    let str = "";
    

    while (n > 0) {
        let rem = n % 26;
        if (rem == 0) {
            str = 'Z' + str;
            n = Math.floor((n / 26)) - 1;
        } else {
            str = String.fromCharCode((rem - 1) + 65) + str;
            n = Math.floor((n / 26));
        }
    }
    return str;
}

for (let i = 1; i <= 100; i++) {
    calcColName(i);
    
    $("#columns").append(`<div class="column-name">${str}</div>`);
    $("#rows").append(`<div class="row-name">${i}</div>`);
}


$("#cells").scroll(function () {
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
});

let cellData = { "Sheet1": {} };
let totalSheets = 1;
let saved = true;
let lastlyAddedSheetNumber = 1;
let selectedSheet = "Sheet1";
let mousemoved = false;
let startCellStored = false;
let startCell;
let endCell;
let defaultProperties = {
    "font-family": "Noto Sans",
    "font-size": 14,
    "text": "",
    "bold": false,
    "italic": false,
    "underlined": false,
    "alignment": "left",
    "color": "#444",
    "bgcolor": "#fff",
    "upStream": [],
    "downStream": []
};
function loadNewSheet() {
    $("#cells").text("");
    for (let i = 1; i <= 100; i++) {
        let row = $('<div class="cell-row"></div>');
        let rowArray = [];
        for (let j = 1; j <= 100; j++) {
            row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);
        }
        // cellData[selectedSheet].push(rowArray);
        $("#cells").append(row);
    }
    addEventsToCells();
    addSheetTabEventListeners();
}

loadNewSheet();

// function addNewSheet() {
//     $(".input-cell").text("");
//     $(".input-cell").css(
//         {
//             "font-family": "Noto Sans",
//             "font-size": 14,
//             "text": "",
//             "bold": false,
//             "italic": false,
//             "underlined": false,
//             "alignment": "left",
//             "color": "#444",
//             "background-color": "#fff"
//         }
//     );
//     addSheetTabEventListeners();
// }

function addEventsToCells() {
    $(".input-cell").dblclick(function () {
        $(this).attr("contenteditable", "true");
        $(this).focus();
    });

    $(".input-cell").blur(function () {
        $(this).attr("contenteditable", "false");
        let [rowId, colId] = findRowCOl(this);
        // cellData[selectedSheet][rowId - 1][colId - 1].text = $(this).text();
        updateCellData("text", $(this).text());
    });

    $(".input-cell").click(function (e) {
        let [rowId, colId] = findRowCOl(this);
        let [topCell, bottomCell, leftCell, rightCell] = getTopBottomLeftRightCell(rowId, colId);


        if ($(this).hasClass("selected") && e.ctrlKey) {
            unselectCell(this, e, topCell, bottomCell, leftCell, rightCell)
        } else {
            selectCell(this, e, topCell, bottomCell, leftCell, rightCell);
        }

    });
    $(".input-cell").mousemove(function (event) {
        event.preventDefault();
        if (event.buttons == 1 && !event.ctrlKey) {
            $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");
            mousemoved = true;
            if (!startCellStored) {
                let [rowId, colId] = findRowCOl(event.target);
                startCell = { rowId: rowId, colId: colId };
                startCellStored = true;
            } else {
                let [rowId, colId] = findRowCOl(event.target);
                endCell = { rowId: rowId, colId: colId };
                selectAllBetweenTheRange(startCell, endCell);
            }
        } else if (event.buttons == 0 && mousemoved) {
            startCellStored = false;
            mousemoved = false;
        }
    })
}
function getTopBottomLeftRightCell(rowId, colId) {
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);
    return [topCell, bottomCell, leftCell, rightCell];
}


function unselectCell(ele, e, topCell, bottomCell, leftCell, rightCell) {
    if (e.ctrlKey && $(ele).attr("contenteditable") == "false") {
        if ($(ele).hasClass("top-selected")) {
            topCell.removeClass("bottom-selected");
        }
        if ($(ele).hasClass("left-selected")) {
            leftCell.removeClass("right-selected");
        }
        if ($(ele).hasClass("right-selected")) {
            rightCell.removeClass("left-selected");
        }
        if ($(ele).hasClass("bottom-selected")) {
            bottomCell.removeClass("top-selected");
        }
        $(ele).removeClass("selected top-selected bottom-selected right-selected left-selected");
    }
}

function selectCell(ele, e, topCell, bottomCell, leftCell, rightCell, mouseSelection) {
    if (e.ctrlKey || mouseSelection) {

        // top selected or not
        let topSelected;
        if (topCell) {
            topSelected = topCell.hasClass("selected");
        }
        // bottom selected or not
        let bottomSelected;
        if (bottomCell) {
            bottomSelected = bottomCell.hasClass("selected");
        }

        // left selected or not
        let leftSelected;
        if (leftCell) {
            leftSelected = leftCell.hasClass("selected");
        }
        // right selected or not
        let rightSelected;
        if (rightCell) {
            rightSelected = rightCell.hasClass("selected");
        }

        if (topSelected) {
            topCell.addClass("bottom-selected");
            $(ele).addClass("top-selected");
        }

        if (leftSelected) {
            leftCell.addClass("right-selected");
            $(ele).addClass("left-selected");
        }

        if (rightSelected) {
            rightCell.addClass("left-selected");
            $(ele).addClass("right-selected");
        }

        if (bottomSelected) {
            bottomCell.addClass("top-selected");
            $(ele).addClass("bottom-selected");
        }
    } else {
        $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");
    }

    $(ele).addClass("selected");
    changeHeader(findRowCOl(ele));
}

function changeHeader([rowId, colId]) {
    let data;
    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {  //agr celldata m h mtlb changes hue h 
        //if c pass mtlb changes hue h 
        data = cellData[selectedSheet][rowId - 1][colId - 1];
    } else {
        data = defaultProperties; //agr ni h toh default prop dedi  fail mtlb koi b changes ni hui 
    }
    $("#font-family").val(data["font-family"]);
    $("#font-family").css("font-family", data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $(".alignment.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");
    addRemoveSelectFromFontStyle(data, "bold");
    addRemoveSelectFromFontStyle(data, "italic");
    addRemoveSelectFromFontStyle(data, "underlined");
    $("#fill-color-icon").css("border-bottom", `4px solid ${data.bgcolor}`);
    $("#text-color-icon").css("border-bottom", `4px solid ${data.color}`);
}

function addRemoveSelectFromFontStyle(data, property) {
    if (data[property]) {
        $(`#${property}`).addClass("selected");
    } else {
        $(`#${property}`).removeClass("selected");
    }
}


function selectAllBetweenTheRange(start, end) {
    for (let i = (start.rowId < end.rowId ? start.rowId : end.rowId); i <= (start.rowId < end.rowId ? end.rowId : start.rowId); i++) {
        for (let j = (start.colId < end.colId ? start.colId : end.colId); j <= (start.colId < end.colId ? end.colId : start.colId); j++) {
            let [topCell, bottomCell, leftCell, rightCell] = getTopBottomLeftRightCell(i, j);
            selectCell($(`#row-${i}-col-${j}`)[0], {}, topCell, bottomCell, leftCell, rightCell, true);
        }
    }
}

$(".menu-selector").change(function (e) {
    let value = $(this).val();
    let key = $(this).attr("id");
    if (key == "font-family") {
        $("#font-family").css(key, value);
    }
    if (!isNaN(value)) {
        value = parseInt(value);
    }
    $(".input-cell.selected").css(key, value);
    // $(".input-cell.selected").each(function (index, data) {
    //     let [rowId, colId] = findRowCOl(data);
    // cellData[selectedSheet][rowId - 1][colId - 1][key] = value;

    // });
    updateCellData(key, value);
})

$(".alignment").click(function (e) {
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    let alignment = $(this).attr("data-type");
    $(".input-cell.selected").css("text-align", alignment);
    // $(".input-cell.selected").each(function (index, data) {
    //     let [rowId, colId] = findRowCOl(data);
    //     cellData[selectedSheet][rowId - 1][colId - 1].alignment = alignment;
    // });
    updateCellData("alignment", alignment);
});

$("#bold").click(function (e) {
    setFontStyle(this, "bold", "font-weight", "bold");
});

$("#italic").click(function (e) {
    setFontStyle(this, "italic", "font-style", "italic");
});

$("#underlined").click(function (e) {
    setFontStyle(this, "underlined", "text-decoration", "underline");
});

function setFontStyle(ele, property, key, value) {
    if ($(ele).hasClass("selected")) {
        $(ele).removeClass("selected");
        $(".input-cell.selected").css(key, "");
        // $(".input-cell.selected").each(function (index, data) {
        //     let [rowId, colId] = findRowCOl(data);
        //     cellData[selectedSheet][rowId - 1][colId - 1][property] = false;
        // });
        updateCellData(property, false);
    } else {
        $(ele).addClass("selected");
        $(".input-cell.selected").css(key, value);
        // $(".input-cell.selected").each(function (index, data) {
        //     let [rowId, colId] = findRowCOl(data);
        //     cellData[selectedSheet][rowId - 1][colId - 1][property] = true;
        // });
        updateCellData(property, true);
    }
}

function updateCellData(property, value) {
    let prevCellData = JSON.stringify(cellData);  //celldata ko store kr lia 
    if (value != defaultProperties[property]) {  //ye check krre h q ki default hi left hti h n agr default h toh delete krni h 
        $(".input-cell.selected").each(function (index, data) {     //hr sleected cell p traverse kra uska cell store krne k lie
            let [rowId, colId] = findRowCOl(data);  //row id colid nikali
            if (cellData[selectedSheet][rowId - 1] == undefined) { //pucha ki rowid exist krti h? ni krti q ki abi k6 change ni hua
                cellData[selectedSheet][rowId - 1] = {};  //row bna q ki exist ni krra 
                cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties }; //col b ni krta hoga islie check ni kia islie col bna dia with default prop 
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
            } else {
                if (cellData[selectedSheet][rowId - 1][colId - 1] == undefined) { //row krti h but col does no exist
                    cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties }; //col b ni krta hoga islie check ni kia islie col bna dia with default prop 
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                } else {
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;  //dono exist krte ho simply lga dia 
                }
            }
        });
    } else {  //left p click kia 
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = findRowCOl(data);
            if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {  //row exist krti h ya ni agr row hi ni exist krta hoga toh undefined aa jaega islie pele use check kia 
                //fr pure cell check kia 
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;  //alignment dedi 
                if (JSON.stringify(cellData[selectedSheet][rowId - 1][colId - 1]) == JSON.stringify(defaultProperties)) {  //dono ko strings m convert kra n compare kr lia 
                    delete cellData[selectedSheet][rowId - 1][colId - 1];  //agr equal h 
                    if (Object.keys(cellData[selectedSheet][rowId - 1]).length == 0) { //col m koi change ni h col exist ni krta 
                        delete cellData[selectedSheet][rowId - 1]; //delete 
                    }
                }
            }
        });
    }
    if (saved && JSON.stringify(cellData) != prevCellData) {
        saved = false;
    }
}


$(".color-pick").colorPick({
    'initialColor': '#TYPECOLOR',
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function () {
        if (this.color != "#TYPECOLOR") {
            if (this.element.attr("id") == "fill-color") {
                $("#fill-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("background-color", this.color);
                // $(".input-cell.selected").each((index, data) => {
                //     let [rowId, colId] = findRowCOl(data);
                //     cellData[selectedSheet][rowId - 1][colId - 1].bgcolor = this.color;
                // });
                updateCellData("bgcolor", this.color);
            } else {
                $("#text-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("color", this.color);
                // $(".input-cell.selected").each((index, data) => {
                //     let [rowId, colId] = findRowCOl(data);
                //     cellData[selectedSheet][rowId - 1][colId - 1].color = this.color;
                // });
                updateCellData("color", this.color);
            }
        }
    }
});

$("#fill-color-icon,#text-color-icon").click(function (e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});

$(".container").click(function (e) {
    $(".sheet-options-modal").remove();
});


function selectSheet(ele) {
    $(".sheet-tab.selected").removeClass("selected");
    $(ele).addClass("selected");
    emptySheet();
    selectedSheet = $(ele).text();
    loadSheet();
}

function emptySheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);  //row nikali 
    for (let i of rowKeys) { //traverse kia 
        let rowId = parseInt(i);  //agr string m hua toh convert 
        let colKeys = Object.keys(data[rowId]);   //colkeys nikali 
        for (let j of colKeys) {  //col m traverse kia 
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text("");
            cell.css({
                "font-family": "Noto Sans",
                "font-size": 14,
                "background-color": "#fff",
                "color": "#444",
                "font-weight": "",
                "font-style": "",
                "text-decoration": "",
                "text-align": "left"
            });
        }
    }
}


function loadSheet() {
    // $("#cells").text("");
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);  //row nikali 
    for (let i of rowKeys) { //traverse kia 
        let rowId = parseInt(i);  //agr string m hua toh convert 
        let colKeys = Object.keys(data[rowId]);   //colkeys nikali 
        for (let j of colKeys) {  //col m traverse kia 
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);  //first cell that have changes
            cell.text(data[rowId][colId].text);  //jha changes hue the unhe reflect kr dia 
            cell.css({
                "font-family": data[rowId][colId]["font-family"],
                "font-size": data[rowId][colId]["font-size"],
                "background-color": data[rowId][colId]["bgcolor"],
                "color": data[rowId][colId].color,
                "font-weight": data[rowId][colId].bold ? "bold" : "",
                "font-style": data[rowId][colId].italic ? "italic" : "",
                "text-decoration": data[rowId][colId].underlined ? "underline" : "",
                "text-align": data[rowId][colId].alignment
            });
        }
    }
}
$(".add-sheet").click(function (e) {
    emptySheet();  //sheet empty kia n nyi sheet bna di pele kia q ki selected hone c pele krna tha 
    totalSheets++;
    lastlyAddedSheetNumber++;
    while (Object.keys(cellData).includes("Sheet" + lastlyAddedSheetNumber)) {
        lastlyAddedSheetNumber++;
    }
    cellData[`Sheet${lastlyAddedSheetNumber}`] = {};
    selectedSheet = `Sheet${lastlyAddedSheetNumber}`;
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheet-tab-container").append(
        `<div class="sheet-tab selected">Sheet${lastlyAddedSheetNumber}</div>`
    );
    $(".sheet-tab.selected")[0].scrollIntoView();
    addSheetTabEventListeners();
    $("#row-1-col-1").click();
    saved = false;  //add m b chnges hote h 

});

function addSheetTabEventListeners() {
    $(".sheet-tab.selected").bind("contextmenu", function (e) {
        e.preventDefault();
        $(".sheet-options-modal").remove();
        let modal = $(`<div class="sheet-options-modal">
                            <div class="option sheet-rename">Rename</div>
                            <div class="option sheet-delete">Delete</div>
                        </div>`);
        $(".container").append(modal);
        $(".sheet-options-modal").css({ "bottom": 0.04 * $(".container").height(), "left": e.pageX });
        $(".sheet-rename").click(function (e) {
            let renameModal = `<div class="sheet-modal-parent">
            <div class="sheet-rename-modal">
                <div class="sheet-modal-title">
                    <span>Rename Sheet</span>
                </div>
                <div class="sheet-modal-input-container">
                    <span class="sheet-modal-input-title">Rename Sheet to:</span>
                    <input class="sheet-modal-input" type="text" />
                </div>
                <div class="sheet-modal-confirmation">
                    <div class="button ok-button">OK</div>
                    <div class="button cancel-button">Cancel</div>
                </div>
            </div>
        </div>`;
            $(".container").append(renameModal);
            $(".cancel-button").click(function (e) {
                $(".sheet-modal-parent").remove();
            });
            $(".ok-button").click(function (e) {
                renameSheet();
            });
            $(".sheet-modal-input").keypress(function (e) {
                if (e.key == "Enter") {
                    renameSheet();
                }
            })
        });

        $(".sheet-delete").click(function (e) {
            let deleteModal = `<div class="sheet-modal-parent">
            <div class="sheet-delete-modal">
                <div class="sheet-modal-title">
                    <span>${$(".sheet-tab.selected").text()}</span>
                </div>
                <div class="sheet-modal-detail-container">
                    <span class="sheet-modal-detail-title">Are you sure?</span>
                </div>
                <div class="sheet-modal-confirmation">
                    <div class="button delete-button">
                        <div class="material-icons delete-icon">delete</div>
                        Delete
                    </div>
                    <div class="button cancel-button">Cancel</div>
                </div>
            </div>
        </div>`;
            $(".container").append(deleteModal);
            $(".cancel-button").click(function (e) {
                $(".sheet-modal-parent").remove();
            });
            $(".delete-button").click(function (e) {
                if (totalSheets > 1) {
                    $(".sheet-modal-parent").remove();
                    let keysArray = Object.keys(cellData);
                    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
                    let currentSelectedSheet = $(".sheet-tab.selected");

                    if (selectedSheetIndex == 0) {
                        selectSheet(currentSelectedSheet.next()[0]);
                    } else {
                        selectSheet(currentSelectedSheet.prev()[0]);  //prev sheet selected thi 
                    }
                    delete cellData[currentSelectedSheet.text()];  //selected sheet ab koi or ho gyi, islie current sheet ko rename kia 
                    currentSelectedSheet.remove();
                    // selectSheet($(".sheet-tab.selected")[0]);
                    totalSheets--;
                    saved = false;
                } else {

                }
            })
        })
        if (!$(this).hasClass("selected")) {
            selectSheet(this);
        }
    });

    $(".sheet-tab.selected").click(function (e) {
        if (!$(this).hasClass("selected")) {
            selectSheet(this);
            $("#row-1-col-1").click();
        }
    });
}

function renameSheet() {
    let newSheetName = $(".sheet-modal-input").val();
    if (newSheetName && !Object.keys(cellData).includes(newSheetName)) {
        let newCellData = {};  //new array bnaya
        for (let i of Object.keys(cellData)) {  //loop chlaya pyrane cell data p
            if (i == selectedSheet) {  //agr equal h 
                newCellData[newSheetName] = cellData[i];
            } else {
                newCellData[i] = cellData[i];
            }
        }
        cellData = newCellData;
        selectedSheet = newSheetName;
        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-modal-parent").remove();
        saved = false;
    } else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(`
            <div class="error"> Sheet Name is not Valid or Sheet already exists! </div>
        `);
    }
}

$(".left-scroller").click(function (e) {

    let keysArray = Object.keys(cellData);
    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
    if (selectedSheetIndex != 0) {
        selectSheet($(".sheet-tab.selected").prev()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView();
});

$(".right-scroller").click(function (e) {
    let keysArray = Object.keys(cellData);
    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
    if (selectedSheetIndex != (keysArray.length - 1)) {
        selectSheet($(".sheet-tab.selected").next()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView();
})

$("#menu-file").click(function (e) {
    let fileModal = $(`<div class="file-modal">
                            <div class="file-options-modal">
                                <div class="close">
                                    <div class="material-icons close-icon">arrow_circle_down</div>
                                    <div>Close</div>
                                </div>
                                <div class="new">
                                    <div class="material-icons new-icon">insert_drive_file</div>
                                    <div>New</div>
                                </div>
                                <div class="open">
                                    <div class="material-icons open-icon">folder_open</div>
                                    <div>Open</div>
                                </div>
                                <div class="save">
                                    <div class="material-icons save-icon">save</div>
                                    <div>Save</div>
                                </div>
                            </div>
                            <div class="file-recent-modal">
                            </div>
                            <div class="file-transparent-modal"></div>
                        </div>`);
    $(".container").append(fileModal);
    fileModal.animate({
        "width": "100vw"
    }, 300);
    $(".close,.file-transparent-modal,.new,.save").click(function (e) {
        fileModal.animate({
            "width": "0vw"
        }, 300);
        setTimeout(() => {
            fileModal.remove();
        }, 299);
    });
    $(".new").click(function (e) {
        if (saved) {
            newFile();
        } else {
            $(".container").append(`<div class="sheet-modal-parent">
                                        <div class="sheet-delete-modal">
                                            <div class="sheet-modal-title">
                                                <span>${$(".title-bar").text()}</span>
                                            </div>
                                            <div class="sheet-modal-detail-container">
                                                <span class="sheet-modal-detail-title">Do you want to save changes?</span>
                                            </div>
                                            <div class="sheet-modal-confirmation">
                                                <div class="button ok-button">
                                                    Save
                                                </div>
                                                <div class="button cancel-button">Cancel</div>
                                            </div>
                                        </div>
                                    </div>`);
            $(".ok-button").click(function (e) {
                $(".sheet-modal-parent").remove();
                saveFile(true);
            });
            $(".cancel-button").click(function (e) {
                $(".sheet-modal-parent").remove();
                newFile();
            })
        }

    });

    $(".save").click(function (e) {
        saveFile();
    });
    $(".open").click(function (e) {
        openFile();
    })
});


function newFile() {
    emptySheet();  //input cell khali css htt jati h n html
    $(".sheet-tab").remove();  //remove
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet1</div>`); //sheet 1 add
    cellData = { "Sheet1": {} }; //bydefault
    selectedSheet = "Sheet1";
    totalSheets = 1;
    lastlyAddedSheetNumber = 1;
    addSheetTabEventListeners();
    $("#row-1-col-1").click();
}

function saveFile(createNewFile) {
    if (!saved) {
        $(".container").append(`<div class="sheet-modal-parent">
                                <div class="sheet-rename-modal">
                                    <div class="sheet-modal-title">
                                        <span>Save File</span>
                                    </div>
                                    <div class="sheet-modal-input-container">
                                        <span class="sheet-modal-input-title">File Name:</span>
                                        <input class="sheet-modal-input" value='${$(".title-bar").text()}' type="text" />
                                    </div>
                                    <div class="sheet-modal-confirmation">
                                        <div class="button ok-button">Save</div>
                                        <div class="button cancel-button">Cancel</div>
                                    </div>
                                </div>
                            </div>`);
        $(".ok-button").click(function (e) {
            let fileName = $(".sheet-modal-input").val();
            if (fileName) {
                let href = `data:application/json,${encodeURIComponent(JSON.stringify(cellData))}`;
                let a = $(`<a href=${href} download="${fileName}.json"></a>`);  //isse download ho jaega 
                $(".container").append(a);
                a[0].click();  //[0] bcz hmne javascript m change kia ye jquery m ni chl ra tha 
                a.remove();
                $(".sheet-modal-parent").remove();
                saved = true;
                if (createNewFile) {
                    newFile();
                }
            }
        });
        $(".cancel-button").click(function (e) {
            $(".sheet-modal-parent").remove();
            if (createNewFile) {
                newFile();
            }
        });
    }
}

function openFile() {
    let inputFile = $(`<input accept="application/json" type="file"/>`);  //input tag file ko accepot krega 
    $(".container").append(inputFile);  //append hua 
    inputFile.click();  //click kia box khula usse file select hoga 
    inputFile.change(function (e) {  //value change ho jati h 
        let file = e.target.files[0];  //file ka object store hta h 
        $(".title-bar").text(file.name.split(".json")[0]);  //name b whi dena h hr file m jo title hoga 
        let reader = new FileReader();  //file reader c read krte h 
        reader.readAsText(file);   //read krega file as a text 
        reader.onload = function () {  //jb reading puri ho jaegi execute this function   
            emptySheet();
            let data = JSON.parse(reader.result);  //data m store kr lia 
            $(".sheet-tab").remove();
            let sheets = Object.keys(cellData);  //sheets aa gyi bhut sari
            for (let i of sheets) {  //sari sheet p traverse krte hue 
                $(".sheet-tab-container").append(`<div class="sheet-tab selected">${i}</div>`)
            }
            addSheetTabEventListeners();
            $(".sheet-tab").removeClass("selected");
            $($(".sheet-tab")[0]).addClass("selected");
            selectedSheet = sheets[0];
            totalSheets = sheets.length;
            lastlyAddedSheetNumber = totalSheets;
            loadSheet();
            inputFile.remove();
        }
    })
}

//cut copy and paste
let clipBoard = { startCell: [], cellData: {} };

$("#cut,#copy").click(function (e) {
    clipBoard.startCell = findRowCOl($(".input-cell.selected")[0]);
    $(".input-cell.selected").each((index, data) => {
        let [rowId, colId] = findRowCOl(data);
        if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
            if (!clipBoard.cellData[rowId]) {
                clipBoard.cellData[rowId] = {};
            }
            clipBoard.cellData[rowId][colId] = { ...cellData[selectedSheet][rowId - 1][colId - 1] };
            if ($(this).text() == "content_cut") {
                delete cellData[selectedSheet][rowId - 1][colId - 1];
                if (Object.keys(cellData[selectedSheet][rowId - 1]).length == 0) {
                    delete cellData[selectedSheet][rowId - 1];
                }
            }
        }
    });
    console.log(cellData);
    console.log(clipBoard);
});

$("#paste").click(function (e) {
    let startCell = findRowCOl($(".input-cell.selected")[0]);
    let rows = Object.keys(clipBoard.cellData);
    for (let i of rows) {
        let cols = Object.keys(clipBoard.cellData[i]);
        for (let j of cols) {
            let rowDistance = parseInt(i) - parseInt(clipBoard.startCell[0]);
            let colDistance = parseInt(j) - parseInt(clipBoard.startCell[1]);
            if (!cellData[selectedSheet][startCell[0] + rowDistance - 1]) {
                cellData[selectedSheet][startCell[0] + rowDistance - 1] = {};
            }
            cellData[selectedSheet][startCell[0] + rowDistance - 1][startCell[1] + colDistance - 1] = { ...clipBoard.cellData[i][j] };
        }
    }
    loadSheet();
})

//formula bar
$("#function-input").blur(function (e) {
    if ($(".input-cell.selected").length > 0) {
        let formula = $(this).text();
        $(".input-cell.selected").each(function (index, data) {
            let tempElements = formula.split(" ");
            // console.log(tempElements);
            let elements = [];
            for (let i of tempElements) {
                if (i.length > 1) {
                    i = i.replace("(", "");
                    i = i.replace(")", "");
                    elements.push(i);
                }
            }
            if (updateStreams(data, elements)) {
                //console.log(cellData);
            } else {
                alert("Formula is invalid!")
            }
        });
    } else {
        alert("Please select a cell first to apply formula!")
    }
});

function updateStreams(ele, elements) {
    let [rowId, colId] = findRowCOl(ele);
    if (!cellData[selectedSheet][rowId - 1]) {
        cellData[selectedSheet][rowId - 1] = {};
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties };
    } else if (!cellData[selectedSheet][rowId - 1][colId - 1]) {
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties };
    }
    cellData[selectedSheet][rowId-1][colId-1].upStream=[];  //pele dalne c pele khali kr dia upstream
    let data = cellData[selectedSheet][rowId - 1][colId - 1];
    for (let i = 0; i < elements.length; i++) {
        if (data.downStream.includes(elements[i])) {
            return false;
        } else {
            if (!data.upStream.includes(elements[i])) {
                data.upStream.push(elements[i]);
            }
        }
    }
    return true;
}

//self cycle detection code
function checkForSelf(rowId, colId, ele) {
    let calRowId;
    let calColId;
    for (let i = 0; i < elements.length; i++) {
        if (!NaN(ele, charAt(i))) {  //no ni h 
            let leftString = ele.substring(0, i);  //0,2 aa  colid
            let rightString = ele.substring(i);  //2 c leke end tk rowid
            calColId=calcColId(leftString);
            calRowId=parseInt(rightString);
            break;
        }

        if(calRowId==rowId&&calColId==colId){
           return true;
        }else{
        
            cellData[selectedSheet][calRowId-1][calColId-1].downStream.push(ele);  //downstream 
            return false;
        }
    }
}

function calcColId(str) {
    let place = str.length - 1
    let total = 0;
    for (let i = 0; i < str.length; i++) {  //from left side traverse
        let charValue = str.CharCodeAt(i) - 64;
        total += Math.pow(26, place) * charValue //26 ki power * char value
        place--;
    }
    return total;
}