// vybe-test: js/string_array_advanced/array_flatmap_filter
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

let arr = ["hello world", "foo bar"];
let words = arr.flatMap(s => s.split(" "));
__check(__line(words.join(",")), "hello,world,foo,bar");
