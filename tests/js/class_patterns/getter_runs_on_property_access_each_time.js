// vybe-test: js/class_patterns/getter_runs_on_property_access_each_time
// origin: languages/js/tests/js/test_class_patterns.rs

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

class Seq {
    constructor() { this.n = 0; }
    get next() { this.n += 1; return this.n; }
}
let s = new Seq();
__check(__line(s.next), "1");
__check(__line(s.next), "2");
