// vybe-test: js/map_set/chain_separate_calls
// origin: languages/js/tests/js/js_map_set_test.rs

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

class B {
            constructor() { this.parts = []; }
            add(s) { this.parts.push(s); return this; }
            build() { return this.parts.join("-"); }
        }
        let b = new B();
        b.add("a");
        b.add("b");
        b.add("c");
        __check(__line(b.build()), "a-b-c");
