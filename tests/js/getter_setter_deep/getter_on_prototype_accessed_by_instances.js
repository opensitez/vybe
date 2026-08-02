// vybe-test: js/getter_setter_deep/getter_on_prototype_accessed_by_instances
// origin: languages/js/tests/js/test_getter_setter_deep.rs

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

function Foo(x) { this.x = x; }
Object.defineProperty(Foo.prototype, "doubled", {
    get() { return this.x * 2; }
});
const a = new Foo(5);
const b = new Foo(10);
__check(__line(a.doubled), "10");
__check(__line(b.doubled), "20");
