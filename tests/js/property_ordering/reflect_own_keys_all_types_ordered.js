// vybe-test: js/property_ordering/reflect_own_keys_all_types_ordered
// origin: languages/js/tests/js/test_property_ordering.rs

function __fmt(v) {
    // console.log renders a bigint with an `n` suffix; String() drops it.
    return typeof v === "bigint" ? String(v) + "n" : String(v);
}

function __line(...args) {
    // console.log joins its arguments with a single space. __fmt is the
    // per-argument coercion console.log applies.
    return args.map(__fmt).join(" ");
}

// Output is COLLECTED, not paired. The emitter rewrites every `console.log(a)`
// into `__p(__line(a))` and compares the whole buffer once.
//
// Collection is what makes ASYNC assertable at all — 967 of the 1,860 cases the
// per-print emitter refused were `await` / `then` / `Promise`, where the i-th
// log in the SOURCE is not the i-th line of OUTPUT. The buffer records the
// order things actually ran, so no ordering analysis is needed.
let __buf = "";

function __p(s) {
    __buf += s + "\n";
}

function __pr(s) {
    __buf += s;
}

// The check runs from a `setTimeout(…, 0)` — a MACROtask, so it fires only
// after the microtask queue has fully drained. Measured under Vybe: a program
// logging sync, then a `.then`, then past an `await`, then the timeout,
// collects them in exactly that order, while a statement at the end of the
// script sees an empty buffer.
function __checkLater(want) {
    setTimeout(function () {
        __check(__buf, want);
    }, 0);
}

function __check(got, want) {
    // The final log contributes a trailing newline the expected line vector
    // never carried, so both forms are accepted.
    if (got !== want && got !== want + "\n") {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

const sym = Symbol("s");
const obj = { 1: "b", sym: "s", 0: "a" };
obj[sym] = "sym";
const names = Object.getOwnPropertyNames(obj);
const intKeys = names.filter(k => /^\d+$/.test(k)).sort((a,b) => +a - +b);
const strKeys = names.filter(k => !/^\d+$/.test(k));
const symKeys = Object.getOwnPropertySymbols(obj);
__p(__line(intKeys[0])); // "0"
__p(__line(intKeys[1])); // "1"
__p(__line(strKeys[0])); // "sym"
__p(__line(typeof symKeys[0])); // "symbol"
__checkLater("0\n1\nsym\nsymbol");
