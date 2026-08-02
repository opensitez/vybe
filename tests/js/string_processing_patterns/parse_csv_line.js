// vybe-test: js/string_processing_patterns/parse_csv_line
// origin: languages/js/tests/js/test_string_processing_patterns.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

function parseCSV(line) {
    return line.split(",").map(s => s.trim());
}
const row = parseCSV("Alice, 30, Engineer");
__check(__line(row[0]), "Alice");
__check(__line(row[1]), "30");
__check(__line(row[2]), "Engineer");
