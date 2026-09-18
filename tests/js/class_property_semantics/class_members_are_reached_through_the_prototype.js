// vybe-test: js/class_property_semantics/class_members_are_reached_through_the_prototype
// origin: ECMA-262 §15.7 — verified against node 26
//
// A class member lives ON `A.prototype` and an instance REACHES it through
// the chain. It is not baked into the instance.
//
// ⛔ THE `delete` ASSERTION IS THE SHARPEST ONE WE HAVE. If the member is
// genuinely on the prototype object, removing it there makes every instance —
// existing and freshly constructed — stop seeing it. If instead it was
// emitted into the instance's own GC struct, `delete` on the prototype cannot
// reach it (you cannot delete a struct field) and the read keeps succeeding.
// One line, unambiguous, and it fails loudly.
//
// It is also outside the family of instruments under test here: `Object.keys`
// and `hasOwnProperty` are themselves suspect when class semantics are wrong,
// so a diagnosis resting on them is not admissible. `delete` is not.
function check(label, got, want) {
    if (got !== want) {
        console.log("FAIL " + label + ": want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed: " + label);
    }
}

class A {
    constructor() { this.x = 1; }
    get g() { return 7; }
    m() { return 2; }
}

// §15.7: class methods and accessors are NON-ENUMERABLE.
const dm = Object.getOwnPropertyDescriptor(A.prototype, "m");
check("descriptor for m exists", typeof dm, "object");
check("m is non-enumerable", dm.enumerable, false);
const dg = Object.getOwnPropertyDescriptor(A.prototype, "g");
check("descriptor for g exists", typeof dg, "object");
check("g is non-enumerable", dg.enumerable, false);
check("keys(A.prototype)", JSON.stringify(Object.keys(A.prototype)), "[]");

const a = new A();
check("a.g before delete", a.g, 7);

// Remove the accessor from the prototype OBJECT.
delete A.prototype.g;
check("a.g after delete", a.g, undefined);
check("keys(a) after delete", JSON.stringify(Object.keys(a)), '["x"]');

// A freshly constructed instance must not see it either — which rules out
// "the instance took a copy at construction time from the live prototype".
const b = new A();
check("new instance after delete", b.g, undefined);
check("keys(new instance)", JSON.stringify(Object.keys(b)), '["x"]');
console.log("ok");
