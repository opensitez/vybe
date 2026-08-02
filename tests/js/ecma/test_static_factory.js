// vybe-test: js/ecma/test_static_factory
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

class User {
            constructor(name, age) {
                this.name = name;
                this.age = age;
            }
            static fromJSON(json) {
                let data = JSON.parse(json);
                return new User(data.name, data.age);
            }
        }
        let u = User.fromJSON('{"name":"Alice","age":30}');
        __check(__line(u.name, u.age), "Alice 30");
