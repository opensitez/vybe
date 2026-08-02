// vybe-test: js/inheritance/test_24_fluent_api_returns_this
// origin: languages/js/tests/js/js_inheritance_test.rs

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
            constructor() { this.parts = ""; }
            add(p) {
                this.parts = this.parts + p;
                return this;
            }
            build() { return this.parts; }
        }
        let result = new Builder().add("a").add("b").add("c").build();
        __check(__line(result), "abc");
