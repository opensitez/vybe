// vybe-test: js/class_private_advanced/public_class_field_with_default_value
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

class Task {
    done = false;
    priority = 1;
    label = "untitled";
}
const t = new Task();
__check(__line(t.done), "false");
__check(__line(t.priority), "1");
__check(__line(t.label), "untitled");
