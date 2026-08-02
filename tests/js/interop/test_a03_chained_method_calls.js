// vybe-test: js/interop/test_a03_chained_method_calls
// origin: languages/js/tests/js/js_interop_test.rs

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

class Builder {
            constructor() { this.parts = []; }
            add(s) { this.parts.push(s); return this; }
            build() { return this.parts.join("-"); }
        }
        let b = new Builder();
        __check(__line(b.add("a").add("b").add("c").build()), "a-b-c");
