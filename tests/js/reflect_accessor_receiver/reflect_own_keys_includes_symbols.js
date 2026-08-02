// vybe-test: js/reflect_accessor_receiver/reflect_own_keys_includes_symbols
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

const s=Symbol("s"); const o={a:1,[s]:2}; const k=Reflect.ownKeys(o); __check(__line(k.includes("a")), "true");__check(__line(k.includes(s)), "true");
