// vybe-test: js/array_higher_order/array_flatmap_flatten_one_level
// origin: languages/js/tests/js/test_array_higher_order.rs

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

const sentences = ["hello world", "foo bar"];
const words = sentences.flatMap(s => s.split(" "));
__check(__line(words.join(",")), "hello,world,foo,bar");
