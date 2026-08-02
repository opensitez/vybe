// vybe-test: js/object_prototype_methods/proto_delete_restores_original
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

const o={}; const orig=Object.getPrototypeOf(o); o.__proto__={z:1}; delete o.__proto__; __check(__line(Object.getPrototypeOf(o)===orig), "true");
