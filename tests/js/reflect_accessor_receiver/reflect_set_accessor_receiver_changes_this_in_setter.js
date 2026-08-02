// vybe-test: js/reflect_accessor_receiver/reflect_set_accessor_receiver_changes_this_in_setter
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

const store={}; const o={set val(v){this.stored=v;}}; Reflect.set(o,"val",42,{stored:undefined,get stored(){return store.v},set stored(x){store.v=x;}}); __check(__line(store.v), "42");
