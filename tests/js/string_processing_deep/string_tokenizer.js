// vybe-test: js/string_processing_deep/string_tokenizer
// origin: languages/js/tests/js/test_string_processing_deep.rs

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

function* tokenize(str) {
    const re = /(\d+\.?\d*)|([a-zA-Z_]\w*)|([+\-*\/()=])/g;
    let m;
    while ((m = re.exec(str)) !== null) {
        if (m[1]) yield { type: "number", value: m[1] };
        else if (m[2]) yield { type: "ident", value: m[2] };
        else yield { type: "op", value: m[3] };
    }
}
const tokens = [...tokenize("x = 3.14 + y")];
__p(__line(tokens.length));
__p(__line(tokens[0].type + ":" + tokens[0].value));
__p(__line(tokens[2].type + ":" + tokens[2].value));
__checkLater("5\nident:x\nnumber:3.14");
