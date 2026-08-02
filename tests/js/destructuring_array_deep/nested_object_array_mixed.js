// vybe-test: js/destructuring_array_deep/nested_object_array_mixed
// origin: languages/js/tests/js/test_destructuring_array_deep.rs

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

const { name, scores: [first, ...rest] } = { name: "Alice", scores: [95, 88, 72] };
__check(__line(name), "Alice");
__check(__line(first), "95");
__check(__line(rest.join(",")), "88,72");
