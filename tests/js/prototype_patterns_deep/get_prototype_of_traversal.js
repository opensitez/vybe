// vybe-test: js/prototype_patterns_deep/get_prototype_of_traversal
// origin: languages/js/tests/js/test_prototype_patterns_deep.rs

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

class A {}
class B extends A {}
class C extends B {}
const c = new C();
const chain = [];
let proto = Object.getPrototypeOf(c);
while (proto !== null) {
    if (proto.constructor) chain.push(proto.constructor.name);
    proto = Object.getPrototypeOf(proto);
}
console.log(chain.includes("C"));
console.log(chain.includes("B"));
console.log(chain.includes("A"));
