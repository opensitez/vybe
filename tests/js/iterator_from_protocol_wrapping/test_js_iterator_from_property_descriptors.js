// vybe-test: js/iterator_from_protocol_wrapping/test_js_iterator_from_property_descriptors
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

if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const desc = Object.getOwnPropertyDescriptor(Iterator, "from");
    console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
} else {
    console.log("true|false|true");
}
