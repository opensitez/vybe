// vybe-test: js/throw_in_loops/for_of_iterator_next_throws
// origin: languages/js/tests/js/test_throw_in_loops.rs

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

let o=[]; try { const bad = { [Symbol.iterator]() { return { next() { throw new Error("next_err"); } }; } }; for (const x of bad) {} } catch(e) { o.push(e.message); } console.log(o.join(","));
