// vybe-test: js/coercion_modern/object_hasown_vs_in
// origin: languages/js/tests/js/test_coercion_modern.rs

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

let parent = { inherited: true };
let child = Object.create(parent);
child.own = true;
__check(__line("inherited" in child), "true");
__check(__line(Object.hasOwn(child, "inherited")), "false");
__check(__line(Object.hasOwn(child, "own")), "true");
