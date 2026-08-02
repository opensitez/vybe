// vybe-test: js/scope_prototype/finalization_registry_register_unregister
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

let fr = new FinalizationRegistry(() => {});
let token = {};
let target = {};
fr.register(target, "held", token);
__check(__line(fr.unregister(token)), "true");
__check(__line(fr.unregister(token)), "false");
