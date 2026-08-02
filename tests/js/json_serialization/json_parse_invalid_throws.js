// vybe-test: js/json_serialization/json_parse_invalid_throws
// origin: languages/js/tests/js/test_json_serialization.rs

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

const invalids = ["undefined", "NaN", "{a:1}", "{'a':1}", "[1,2,]"];
let count = 0;
for (const s of invalids) {
    try { JSON.parse(s); }
    catch { count++; }
}
console.log(count);
