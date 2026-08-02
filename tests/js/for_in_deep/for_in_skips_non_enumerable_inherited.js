// vybe-test: js/for_in_deep/for_in_skips_non_enumerable_inherited
// origin: languages/js/tests/js/test_for_in_deep.rs

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

function Parent() {}
Object.defineProperty(Parent.prototype, "hidden", { value: 1, enumerable: false });
Parent.prototype.visible = 2;
const obj = new Parent();
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("hidden"));
console.log(keys.includes("visible"));
