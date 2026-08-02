// vybe-test: js/reflect_accessor_receiver/reflect_set_invokes_setter_with_receiver
// origin: languages/js/tests/js/test_reflect_accessor_receiver.rs

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

const target={}; const recv={v:0}; const o={set x(val){this.v=val;}}; Object.setPrototypeOf(o,target); Reflect.set(o,"x",5,recv); __check(__line(recv.v), "5");
