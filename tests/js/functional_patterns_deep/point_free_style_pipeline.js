// vybe-test: js/functional_patterns_deep/point_free_style_pipeline
// origin: languages/js/tests/js/test_functional_patterns_deep.rs

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

const words = ["Hello", "World", "Foo", "Bar"];
const process = arr => arr
    .map(s => s.toLowerCase())
    .filter(s => s.length > 3)
    .sort()
    .join(",");
__check(__line(process(words)), "hello,world");
