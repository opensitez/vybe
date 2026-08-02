// vybe-test: js/object_advanced/object_getprototypeof_object_create_chain
// origin: languages/js/tests/js/test_object_advanced.rs

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

let proto = { kind: "base" };
let obj = Object.create(proto);
__check(__line(Object.getPrototypeOf(obj) === proto), "true");
__check(__line(obj.kind), "base");
