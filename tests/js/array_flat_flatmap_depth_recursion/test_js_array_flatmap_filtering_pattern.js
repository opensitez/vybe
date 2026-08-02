// vybe-test: js/array_flat_flatmap_depth_recursion/test_js_array_flatmap_filtering_pattern
// origin: languages/js/tests/js/test_js_array_flat_flatmap_depth_recursion.rs

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

const sentence = ["Hello World", "JS Engine"];
const words = sentence.flatMap(s => s.split(" "));
__check(__line(words.join(",")), "Hello,World,JS,Engine");
