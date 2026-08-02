// vybe-test: js/class_private_advanced/static_init_block_runs_once_at_class_definition
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

let ran = 0;
class Setup {
    static {
        ran++;
    }
}
__check(__line(ran), "1");
const a = new Setup();
const b = new Setup();
__check(__line(ran), "1");
