// vybe-test: js/variadic_bug/variadic_instance_method_call_packs_rest
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
            call(prefix, ...parts) {
                return prefix + ":" + parts.join(",");
            }
        }
        const g = new Greeter();
        __check(__line(g.call("head", "a", "b", "c")), "head:a,b,c");
