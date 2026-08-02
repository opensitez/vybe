// vybe-test: js/misc_es_features/string_at_with_array_map
// origin: languages/js/tests/js/test_misc_es_features.rs

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

const words = ["hello", "world"];
const firsts = words.map(w => w.at(0));
__check(__line(firsts.join(",")), "h,w");
