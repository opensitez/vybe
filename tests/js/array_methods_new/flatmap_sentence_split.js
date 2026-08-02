// vybe-test: js/array_methods_new/flatmap_sentence_split
// origin: languages/js/tests/js/test_array_methods_new.rs

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

const sentences = ["Hello World", "Foo Bar"];
const words = sentences.flatMap(s => s.split(" "));
__check(__line(words.join(",")), "Hello,World,Foo,Bar");
