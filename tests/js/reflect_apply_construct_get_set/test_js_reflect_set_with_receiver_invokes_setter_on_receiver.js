// vybe-test: js/reflect_apply_construct_get_set/test_js_reflect_set_with_receiver_invokes_setter_on_receiver
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set.rs

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

const proto = {
    set score(v) { this._score = v * 2; }
};
const receiver = {};
const success = Reflect.set(proto, "score", 50, receiver);
__check(__line(success + "|" + receiver._score), "true|100");
