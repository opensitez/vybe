// vybe-test: js/symbol_wellknown/symbol_search_custom_search
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

const searcher = {
  [Symbol.search](str) { return str.indexOf("x"); }
};
__check(__line("abxcd".search(searcher)), "2");
