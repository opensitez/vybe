// vybe-test: js/object_prototype_methods/proto_getter_matches_getprototypeof
// origin: languages/js/tests/js/test_object_prototype_methods.rs

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

const p={a:1}; const o=Object.create(p); __check(__line(o.__proto__===p), "true"); __check(__line(Object.getPrototypeOf(o)===p), "true");
