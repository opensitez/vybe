// vybe-test: js/object_destructuring_default_values_aliases/test_js_object_destructuring_assignment_existing_variables
// origin: languages/js/tests/js/test_js_object_destructuring_default_values_aliases.rs

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

let x, y;
({ x, y } = { x: 1, y: 2 });
__check(__line(x + "|" + y), "1|2");
