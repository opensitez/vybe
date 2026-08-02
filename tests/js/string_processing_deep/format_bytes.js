// vybe-test: js/string_processing_deep/format_bytes
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

function formatBytes(bytes) {
    const units = ["B","KB","MB","GB","TB"];
    let i = 0;
    while (bytes >= 1024 && i < units.length - 1) { bytes /= 1024; i++; }
    return bytes.toFixed(i === 0 ? 0 : 2) + " " + units[i];
}
console.log(formatBytes(0));
console.log(formatBytes(1024));
console.log(formatBytes(1024 * 1024));
console.log(formatBytes(1500));
