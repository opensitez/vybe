// vybe-test: js/arrow_functions/arrow_property_this_is_outer_scope
// origin: languages/js/tests/js/test_arrow_functions.rs

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

const o={v:1,f:()=>this}; __check(__line(o.f()===o.f()), "true"); __check(__line(o.f()!==o), "true");
