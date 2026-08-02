// vybe-test: js/object_advanced/for_in_own_vs_inherited
// origin: languages/js/tests/js/test_object_advanced.rs

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

let parent = { a: 1 };
let child = Object.create(parent);
child.b = 2;
let own = [];
let all = [];
for (let key in child) {
    all.push(key);
    if (child.hasOwnProperty(key)) own.push(key);
}
console.log(own.join(","));
console.log(all.sort().join(","));
