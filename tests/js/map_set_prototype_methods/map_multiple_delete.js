// vybe-test: js/map_set_prototype_methods/map_multiple_delete
// origin: languages/js/tests/js/test_map_set_prototype_methods.rs

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

const m=new Map([["a",1],["b",2],["c",3]]); m.delete("b"); m.delete("a"); __check(__line(m.size), "1"); __check(__line(m.has("c")), "true");
