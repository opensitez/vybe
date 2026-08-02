// vybe-test: js/function_deep/new_target_allows_dual_use_constructor
// origin: languages/js/tests/js/test_function_deep.rs

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

function Greeter(name) {
    if (new.target === undefined) return new Greeter(name);
    this.name = name;
}
const a = new Greeter("Alice");
const b = Greeter("Bob"); // works without new
__check(__line(a.name), "Alice");
__check(__line(b.name), "Bob");
