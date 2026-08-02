// vybe-test: js/iterator_protocol/infinite_iterator_with_take_via_break
// origin: languages/js/tests/js/test_iterator_protocol.rs

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

function* naturals() {
    let n = 1;
    while (true) yield n++;
}
const first5 = [];
for (const n of naturals()) {
    if (n > 5) break;
    first5.push(n);
}
console.log(first5.join(","));
