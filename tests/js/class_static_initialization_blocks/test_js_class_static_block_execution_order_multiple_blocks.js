// vybe-test: js/class_static_initialization_blocks/test_js_class_static_block_execution_order_multiple_blocks
// origin: languages/js/tests/js/test_js_class_static_initialization_blocks.rs

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
class OrderTest {
    static field1 = (() => { log.push("Field 1"); return 1; })();
    static {
        log.push("Static Block 1");
    }
    static field2 = (() => { log.push("Field 2"); return 2; })();
    static {
        log.push("Static Block 2");
    }
}
__check(__line(log.join("->")), "Field 1->Static Block 1->Field 2->Static Block 2");
