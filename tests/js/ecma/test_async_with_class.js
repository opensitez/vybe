// vybe-test: js/ecma/test_async_with_class
// origin: languages/js/tests/js/js_ecma_test.rs

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

class UserService {
            constructor(name) {
                this.name = name;
            }
            async greet() {
                return "Hello, " + this.name + "!";
            }
        }
        
        let svc = new UserService("Bob");
        console.log(await svc.greet());
