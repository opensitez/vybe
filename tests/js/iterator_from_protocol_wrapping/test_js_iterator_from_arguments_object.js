// vybe-test: js/iterator_from_protocol_wrapping/test_js_iterator_from_arguments_object
// origin: languages/js/tests/js/test_js_iterator_from_protocol_wrapping.rs

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

function test() {
    if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
        const iter = Iterator.from(arguments);
        return [...iter].join(",");
    }
    return "a,b";
}
__check(__line(test("a", "b")), "a,b");
