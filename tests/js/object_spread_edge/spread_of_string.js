// vybe-test: js/object_spread_edge/spread_of_string
// origin: languages/js/tests/js/test_object_spread_edge.rs

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

const str = "abc";
const chars = {};
for (let i = 0; i < str.length; i++) chars[i] = str[i];
console.log(chars[0]);
console.log(chars[1]);
console.log(chars[2]);
