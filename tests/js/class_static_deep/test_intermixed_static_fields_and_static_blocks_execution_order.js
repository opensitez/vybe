// vybe-test: js/class_static_deep/test_intermixed_static_fields_and_static_blocks_execution_order
// origin: languages/js/tests/js/test_class_static_deep.rs

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

const log = [];
class C {
    static a = (log.push("fieldA"), 1);
    static {
        log.push("block1");
    }
    static b = (log.push("fieldB"), 2);
    static {
        log.push("block2");
    }
}
__check(__line(log.join(",")), "fieldA,block1,fieldB,block2");
