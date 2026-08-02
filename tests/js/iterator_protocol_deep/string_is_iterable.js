// vybe-test: js/iterator_protocol_deep/string_is_iterable
// origin: languages/js/tests/js/test_iterator_protocol_deep.rs

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

const chars = [];
for (const c of "hello") chars.push(c);
console.log(chars.join("-"));
