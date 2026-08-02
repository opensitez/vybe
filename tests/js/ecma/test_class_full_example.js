// vybe-test: js/ecma/test_class_full_example
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

class EventEmitter {
            constructor() {
                this._handlers = {};
            }
            on(event, handler) {
                if (!this._handlers[event]) {
                    this._handlers[event] = [];
                }
                this._handlers[event].push(handler);
            }
            emit(event, data) {
                let handlers = this._handlers[event];
                if (handlers) {
                    handlers.forEach((h) => h(data));
                }
            }
        }
        
        let emitter = new EventEmitter();
        let log = [];
        emitter.on("greet", (name) => { log.push("Hello " + name); });
        emitter.on("greet", (name) => { log.push("Hi " + name); });
        emitter.emit("greet", "Alice");
        console.log(log.join(", "));
