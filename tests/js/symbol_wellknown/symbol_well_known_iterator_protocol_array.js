// vybe-test: js/symbol_wellknown/symbol_well_known_iterator_protocol_array
// origin: languages/js/tests/js/test_symbol_wellknown.rs

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

const iter = [1, 2, 3][Symbol.iterator]();
__check(__line(iter.next().value), "1");
__check(__line(iter.next().value), "2");
__check(__line(iter.next().done), "false");
