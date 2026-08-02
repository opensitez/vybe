// vybe-test: js/explicit_resource_management/using_disposes_in_lifo_order
// origin: languages/js/tests/js/test_explicit_resource_management.rs

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

const order = [];
{
    using r1 = { [Symbol.dispose]() { order.push(1); } };
    using r2 = { [Symbol.dispose]() { order.push(2); } };
    using r3 = { [Symbol.dispose]() { order.push(3); } };
}
__check(__line(order.join(",")), "3,2,1");
