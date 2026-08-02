// vybe-test: js/object_prototype_methods/proto_set_shadows_inherited
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

const p={x:1}; const o=Object.create(p); o.__proto__={x:2}; __check(__line(o.x), "2");
