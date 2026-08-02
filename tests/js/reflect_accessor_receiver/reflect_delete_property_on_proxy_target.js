// vybe-test: js/reflect_accessor_receiver/reflect_delete_property_on_proxy_target
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

const t={a:1}; const o=new Proxy(t,{}); __check(__line(Reflect.deleteProperty(o,"a")), "true");__check(__line("a" in t), "false");
