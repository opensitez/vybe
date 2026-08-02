// vybe-test: js/nullish_coalescing_and_optional_chaining_combinations/test_js_optional_chaining_short_circuits_entire_chain
// origin: languages/js/tests/js/test_js_nullish_coalescing_and_optional_chaining_combinations.rs

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

let sideEffectCount = 0;
const fn = () => { sideEffectCount++; return "data"; };
const obj = null;

const res = obj?.prop[fn()];
__check(__line((res === undefined) + "|SideEffects=" + sideEffectCount), "true|SideEffects=0");
