// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_super_constructor_initialization_order
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

const log = [];
class Base {
    constructor() {
        log.push("Base Ctor");
    }
}
class Derived extends Base {
    #field = (() => { log.push("Init Field"); return 10; })();
    constructor() {
        log.push("Before Super");
        super();
        log.push("After Super");
    }
}
new Derived();
__check(__line(log.join("->")), "Before Super->Base Ctor->Init Field->After Super");
