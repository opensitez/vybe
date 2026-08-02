// vybe-test: js/string_processing_deep/csv_serializer
// origin: languages/js/tests/js/test_string_processing_deep.rs

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

function toCSV(data, headers) {
    const escape = v => /[,"\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v);
    const rows = data.map(row => headers.map(h => escape(row[h])).join(","));
    return [headers.join(","), ...rows].join("\n");
}
const data = [
    { name: "Alice", age: 30, city: "New York" },
    { name: "Bob", age: 25, city: "Los Angeles" },
];
const csv = toCSV(data, ["name", "age", "city"]);
const lines = csv.split("\n");
console.log(lines[0]);
console.log(lines[1]);
