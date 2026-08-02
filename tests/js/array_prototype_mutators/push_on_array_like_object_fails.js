// vybe-test: js/array_prototype_mutators/push_on_array_like_object_fails
// origin: languages/js/tests/js/test_array_prototype_mutators.rs

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

const o={length:0, push:Array.prototype.push}; o.push(1); __check(__line(o[0]), "1");__check(__line(o.length), "1");
