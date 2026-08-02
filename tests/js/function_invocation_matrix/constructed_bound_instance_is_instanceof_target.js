// vybe-test: js/function_invocation_matrix/constructed_bound_instance_is_instanceof_target
// origin: languages/js/tests/js/test_function_invocation_matrix.rs

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

function Animal(kind) {
    this.kind = kind;
}
const Dog = Animal.bind(null, "dog");
const d = new Dog();
__check(__line(d instanceof Animal), "true");
__check(__line(d.kind), "dog");
