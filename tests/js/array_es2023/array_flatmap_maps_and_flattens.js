// vybe-test: js/array_es2023/array_flatmap_maps_and_flattens
// origin: languages/js/tests/js/test_array_es2023.rs

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

const arr = ["hello world", "foo bar baz"];
const words = arr.flatMap(s => s.split(" "));
__check(__line(words.length), "5");
__check(__line(words.join(",")), "hello,world,foo,bar,baz");
