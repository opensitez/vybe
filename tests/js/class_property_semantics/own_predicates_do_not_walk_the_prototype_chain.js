// vybe-test: js/class_property_semantics/own_predicates_do_not_walk_the_prototype_chain
// origin: ECMA-262 §20.1.3.2 / §10.1.11 — verified against node 26
//
// A class accessor lives on the PROTOTYPE (§15.7). The own-property
// predicates must therefore report it on `A.prototype` and NOT on an
// instance. Both of these have to hold at once, and one of them failing
// while the other passes is the signature of an own-predicate that walks
// the chain and attributes what it finds to the receiver.
//
// ⛔ THE CORPUS COULD NOT SEE THIS. Every existing enumeration suite puts a
// plain object or `Object.defineProperty` on the left-hand side; a class is a
// different emission path, so all of them stayed green while these answers
// were wrong.
function check(label, got, want) {
    if (got !== want) {
        console.log("FAIL " + label + ": want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed: " + label);
    }
}

class A {
    constructor() { this.x = 1; }
    get g() { return 7; }
    set s(v) {}
    m() { return 2; }
}
const a = new A();
const H = Object.prototype.hasOwnProperty;

// The instance owns its data property and nothing else.
check("keys(a)", JSON.stringify(Object.keys(a)), '["x"]');
check("gOPN(a)", JSON.stringify(Object.getOwnPropertyNames(a)), '["x"]');
check("hasOwn(a,'g')", H.call(a, "g"), false);
check("hasOwn(a,'s')", H.call(a, "s"), false);
check("hasOwn(a,'m')", H.call(a, "m"), false);

// The accessor really is on the prototype — so the answers above are about
// OWNERSHIP, not about the accessor being missing.
check("hasOwn(proto,'g')", H.call(A.prototype, "g"), true);
check("a.g still reads through the chain", a.g, 7);

// `Object.keys` is own+enumerable, `getOwnPropertyNames` is own including
// non-enumerables — so keys must always be a SUBSET.
const kNames = Object.getOwnPropertyNames(a);
for (const k of Object.keys(a)) {
    check("keys(a) subset of gOPN(a): " + k, kNames.indexOf(k) >= 0, true);
}
console.log("ok");
