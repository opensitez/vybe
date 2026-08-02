// vybe-test: js/number_boolean_prototype/number_toprecision_range_error_high
// origin: languages/js/tests/js/test_number_boolean_prototype.rs

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

try{console.log((1).toPrecision(101));}catch(e){console.log(e instanceof RangeError);}
