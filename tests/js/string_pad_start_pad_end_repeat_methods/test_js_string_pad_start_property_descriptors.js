// vybe-test: js/string_pad_start_pad_end_repeat_methods/test_js_string_pad_start_property_descriptors
// origin: languages/js/tests/js/test_js_string_pad_start_pad_end_repeat_methods.rs

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

const dStart = Object.getOwnPropertyDescriptor(String.prototype, "padStart");
const dRepeat = Object.getOwnPropertyDescriptor(String.prototype, "repeat");
__check(__line(`${dStart.writable}:${dStart.configurable}:${dRepeat.writable}:${dRepeat.configurable}`), "true:true:true:true");
