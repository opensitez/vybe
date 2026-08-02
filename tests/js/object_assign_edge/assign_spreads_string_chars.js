// vybe-test: js/object_assign_edge/assign_spreads_string_chars
// origin: languages/js/tests/js/test_object_assign_edge.rs

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
const result = {};
for (let i = 0; i < str.length; i++) result[i] = str[i];
console.log(result[0]);
console.log(result[1]);
console.log(result[2]);
