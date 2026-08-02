// vybe-test: js/destructuring_advanced/destructure_object_in_for_of
// origin: languages/js/tests/js/test_destructuring_advanced.rs

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

const people = [{ name: "X", age: 1 }, { name: "Y", age: 2 }];
const names = [];
for (const { name } of people) names.push(name);
console.log(names.join(","));
