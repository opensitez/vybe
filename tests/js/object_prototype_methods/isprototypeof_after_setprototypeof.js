// vybe-test: js/object_prototype_methods/isprototypeof_after_setprototypeof
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

const a={}; const b={}; const c={}; Object.setPrototypeOf(c,a); __check(__line(a.isPrototypeOf(c)), "true"); Object.setPrototypeOf(c,b); __check(__line(a.isPrototypeOf(c)), "false"); __check(__line(b.isPrototypeOf(c)), "true");
