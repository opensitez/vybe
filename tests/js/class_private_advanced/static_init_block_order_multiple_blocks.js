// vybe-test: js/class_private_advanced/static_init_block_order_multiple_blocks
// origin: languages/js/tests/js/test_class_private_advanced.rs

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

class Ordered {
    static log = [];
    static { Ordered.log.push("first"); }
    static { Ordered.log.push("second"); }
    static { Ordered.log.push("third"); }
}
__check(__line(Ordered.log.join(",")), "first,second,third");
