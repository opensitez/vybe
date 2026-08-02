// vybe-test: js/generator_state_machines/csv_parser_generator
// origin: languages/js/tests/js/test_generator_state_machines.rs

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

function* parseCSV(text) {
    const lines = text.split("\n").filter(Boolean);
    for (const line of lines) {
        yield line.split(",").map(s => s.trim());
    }
}
const csv = "Alice,30,Engineer\nBob,25,Designer\nCharlie,35,Manager";
const rows = [...parseCSV(csv)];
console.log(rows.length);
console.log(rows[0][0]);
console.log(rows[1][2]);
