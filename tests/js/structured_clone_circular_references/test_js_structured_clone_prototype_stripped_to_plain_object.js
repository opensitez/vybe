// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_prototype_stripped_to_plain_object
// origin: languages/js/tests/js/test_js_structured_clone_circular_references.rs

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

class CustomClass {
    constructor() { this.x = 10; }
}
const inst = new CustomClass();
const clone = structuredClone(inst);
__check(__line((clone.x === 10) + "|isCustom=" + (clone instanceof CustomClass) + "|isObject=" + (clone.constructor === Object)), "true|isCustom=false|isObject=true");
