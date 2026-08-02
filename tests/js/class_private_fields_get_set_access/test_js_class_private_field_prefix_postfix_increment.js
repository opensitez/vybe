// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_prefix_postfix_increment
// origin: languages/js/tests/js/test_js_class_private_fields_get_set_access.rs

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

class Counter {
    #x = 10;
    inc() {
        const pre = ++this.#x;
        const post = this.#x++;
        return `${pre}:${post}:${this.#x}`;
    }
}
__check(__line(new Counter().inc()), "11:11:12");
