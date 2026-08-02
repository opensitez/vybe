// vybe-test: js/property_enumeration/array_indices_sorted_as_integers
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const arr = {};
arr[100] = "c";
arr[2] = "b";
arr[1] = "a";
arr.extra = "e";
const keys = Object.keys(arr);
const intKeys = keys.filter(k => /^\d+$/.test(k)).sort((a,b) => +a - +b);
const strKeys = keys.filter(k => !/^\d+$/.test(k));
__check(__line([...intKeys, ...strKeys].join(",")), "1,2,100,extra");
