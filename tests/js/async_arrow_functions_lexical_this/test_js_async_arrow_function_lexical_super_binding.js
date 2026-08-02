// vybe-test: js/async_arrow_functions_lexical_this/test_js_async_arrow_function_lexical_super_binding
// origin: languages/js/tests/js/test_js_async_arrow_functions_lexical_this.rs

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

class Base {
    async getName() { return "BaseName"; }
}
class Sub extends Base {
    async getName() {
        const getSuper = async () => await super.getName();
        return (await getSuper()) + "_Extended";
    }
}
new Sub().getName().then(res => console.log(res));
