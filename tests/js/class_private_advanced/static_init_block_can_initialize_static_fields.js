// vybe-test: js/class_private_advanced/static_init_block_can_initialize_static_fields
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

class Constants {
    static PI;
    static E;
    static {
        Constants.PI = 3.14159;
        Constants.E = 2.71828;
    }
}
__check(__line(Constants.PI), "3.14159");
__check(__line(Constants.E), "2.71828");
