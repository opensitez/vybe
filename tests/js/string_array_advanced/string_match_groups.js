// vybe-test: js/string_array_advanced/string_match_groups
// origin: languages/js/tests/js/test_string_array_advanced.rs

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

let s = "2024-01-15";
let m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
__check(__line(m[1]), "2024");
__check(__line(m[2]), "01");
__check(__line(m[3]), "15");
