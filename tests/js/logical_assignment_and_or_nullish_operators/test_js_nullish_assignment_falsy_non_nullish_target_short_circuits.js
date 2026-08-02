// vybe-test: js/logical_assignment_and_or_nullish_operators/test_js_nullish_assignment_falsy_non_nullish_target_short_circuits
// origin: languages/js/tests/js/test_js_logical_assignment_and_or_nullish_operators.rs

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

let a = 0;
let b = "";
let c = false;

a ??= 99;
b ??= "default";
c ??= true;
__check(__line(`${a}:${b}:${c}`), "0::false");
