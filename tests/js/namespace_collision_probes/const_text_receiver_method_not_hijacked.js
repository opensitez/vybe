// vybe-test: js/namespace_collision_probes/const_text_receiver_method_not_hijacked
// origin: languages/js/tests/js/test_namespace_collision_probes.rs

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

const text = "a1b2"; __check(__line([...text.matchAll(/\d/g)].length), "2");
