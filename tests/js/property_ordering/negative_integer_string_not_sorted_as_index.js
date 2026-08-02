// vybe-test: js/property_ordering/negative_integer_string_not_sorted_as_index
// origin: languages/js/tests/js/test_property_ordering.rs

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

const obj = { "-1": "neg", 0: "zero", a: "a" };
const keys = Object.keys(obj);
const intKeys = keys.filter(k => /^\d+$/.test(k)).sort((a,b) => +a - +b);
const strKeys = keys.filter(k => !/^\d+$/.test(k));
const sorted = [...intKeys, ...strKeys];
__check(__line(sorted[0]), "0"); // 0 (integer index first)
__check(__line(sorted.includes("-1")), "true");
