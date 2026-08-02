// vybe-test: js/closures_functional/custom_flat_map
// origin: languages/js/tests/js/test_closures_functional.rs

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

let sentences = ["Hello World", "Foo Bar Baz"];
let words = sentences.flatMap(s => s.split(" "));
__check(__line(words.join(",")), "Hello,World,Foo,Bar,Baz");
