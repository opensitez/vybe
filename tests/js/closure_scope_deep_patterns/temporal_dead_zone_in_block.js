// vybe-test: js/closure_scope_deep_patterns/temporal_dead_zone_in_block
// origin: languages/js/tests/js/test_closure_scope_deep_patterns.rs

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

let result = "before";
{
    // let x is not initialized yet (TDZ if accessed here)
    result = "in block";
    let x = 10;
    result = "after let: " + x;
}
__check(__line(result), "after let: 10");
