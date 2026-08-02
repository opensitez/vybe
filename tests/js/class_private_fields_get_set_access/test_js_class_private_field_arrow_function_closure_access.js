// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_arrow_function_closure_access
// origin: languages/js/tests/js/test_js_class_private_fields_get_set_access.rs

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

class Service {
    #token = "SECRET_TOKEN";
    getFetcher() {
        return () => `Bearer ${this.#token}`;
    }
}
const s = new Service();
const fetcher = s.getFetcher();
__check(__line(fetcher()), "Bearer SECRET_TOKEN");
