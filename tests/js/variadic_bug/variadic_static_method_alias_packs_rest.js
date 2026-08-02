// vybe-test: js/variadic_bug/variadic_static_method_alias_packs_rest
// origin: languages/js/tests/js/test_variadic_bug.rs

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

class Greeter {
            static call(prefix, ...parts) {
                return prefix + ":" + parts.join(",");
            }
        }
        const alias = Greeter.call;
        __check(__line(alias("head", "a", "b", "c")), "head:a,b,c");
