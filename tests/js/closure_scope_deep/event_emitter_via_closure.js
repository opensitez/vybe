// vybe-test: js/closure_scope_deep/event_emitter_via_closure
// origin: languages/js/tests/js/test_closure_scope_deep.rs

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

function createEmitter() {
    const listeners = new Map();
    return {
        on(event, fn) {
            if (!listeners.has(event)) listeners.set(event, []);
            listeners.get(event).push(fn);
        },
        emit(event, data) {
            listeners.get(event)?.forEach(fn => fn(data));
        }
    };
}
const em = createEmitter();
const log = [];
em.on("data", v => log.push("a:" + v));
em.on("data", v => log.push("b:" + v));
em.emit("data", 42);
console.log(log.join(","));
