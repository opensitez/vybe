// vybe-test: js/reflect_accessor_receiver/reflect_set_on_proxy_forwards_to_target
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

const t={}; const o=new Proxy(t,{}); Reflect.set(o,"a",1); __check(__line(t.a), "1");
