// vybe-test: js/property_ordering/keys_integer_indices_sorted_first
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

const obj = { b: 2, 0: "zero", a: 1, 2: "two", 1: "one" };
const keys = Object.keys(obj);
const intKeys = keys.filter(k => /^\d+$/.test(k)).sort((a,b) => +a - +b);
const strKeys = keys.filter(k => !/^\d+$/.test(k));
__check(__line([...intKeys, ...strKeys].join(",")), "0,1,2,b,a");
