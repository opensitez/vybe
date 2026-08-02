// vybe-test: js/function_prototype_deep/bind_extracted_instance_method_preserves_this
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

const counter = { n: 0, tick() { this.n++; return this.n; } }; const tick = counter.tick.bind(counter); __check(__line(tick()), "1");
