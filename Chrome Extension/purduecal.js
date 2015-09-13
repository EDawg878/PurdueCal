var table = $("table[summary^='Display course details']");
var end = table.find("tr").length - 4;

table.find("td").each(function() {
    var col = $(this);
    if (row_index == end) {
        return false;
    } else if (col.html() == "&nbsp;") {
        var row_index = col.parent().index();
        var col_index = col.index();
        var query_above = "tr:eq(" + (row_index - 1) + ") td:eq(" + col_index+ ")";
        var query_right = "tr:eq(" + row_index + ") td:eq(" + (col_index + 1) + ")";
        var col_above = $(query_above, table);
        var col_right = $(query_right, table);
        if (col_right.html() == '<abbr title="To Be Announced">TBA</abbr>') {
            col.html(col_right.html());
        } else {
            col.html(col_above.html());
        }
    }
});

// TODO make this change not show to users
table.find("tr:last").remove();

// via http://stackoverflow.com/a/19305829
function download(filename, content) {
    var a = document.createElement('a');
    var blob = new Blob([ content ], {type : "text/plain;charset=UTF-8"});
    a.href = window.URL.createObjectURL(blob);
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    delete a;
}

var json = JSON.stringify(table.tableToJSON());
download("MyPurdueSchedule", json);