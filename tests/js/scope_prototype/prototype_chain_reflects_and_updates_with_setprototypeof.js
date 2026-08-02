// vybe-test: js/scope_prototype/prototype_chain_reflects_and_updates_with_setprototypeof
// origin: languages/js/tests/js/test_scope_prototype.rs

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

const baseA = { mode: "A" };
const baseB = { mode: "B" };
const obj = Object.create(baseA);

__check(__line(Object.getPrototypeOf(obj) === baseA), "true");
__check(__line(obj.mode), "A");

Object.setPrototypeOf(obj, baseB);
__check(__line(Object.getPrototypeOf(obj) === baseB), "true");
__check(__line(obj.mode), "B");

Object.setPrototypeOf(obj, null);
__check(__line(Object.getPrototypeOf(obj) === null), "true");
__check(__line("mode" in obj), "false");
