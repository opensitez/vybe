// vybe-test: js/scope_prototype/block_scope_shadow_does_not_modify_outer_binding
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let value = 1;
{
    let value = 2;
    value += 3;
    __check(__line(value), "5");
}
__check(__line(value), "1");
