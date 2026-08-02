// vybe-test: js/scope_prototype/prototype_property_falls_back_when_instance_missing
// origin: languages/js/tests/js/test_scope_prototype.rs

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

function Person() {}
Person.prototype.role = "user";
let p = new Person();
__check(__line(p.role), "user");
p.role = "admin";
__check(__line(p.role), "admin");
delete p.role;
__check(__line(p.role), "user");
