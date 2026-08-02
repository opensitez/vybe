// vybe-test: js/structured_clone_error_types_serialization/test_js_structured_clone_custom_error_class_prototype_fallback
// origin: languages/js/tests/js/test_js_structured_clone_error_types_serialization.rs

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

class CustomAppError extends Error {
    constructor(msg) {
        super(msg);
        this.name = "CustomAppError";
    }
}
const err = new CustomAppError("AppFailed");
const clone = structuredClone(err);

__check(__line(clone.name + "|" + (clone instanceof Error) + "|isCustom=" + (clone instanceof CustomAppError)), "CustomAppError|true|isCustom=false");
