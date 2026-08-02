// vybe-test: js/reactive_patterns/signal_reactivity
// origin: languages/js/tests/js/test_reactive_patterns.rs

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

function createSignal(init) {
    let value = init;
    const subscribers = new Set();
    const get = () => value;
    const set = (v) => { value = v; subscribers.forEach(fn => fn(v)); };
    const subscribe = fn => { subscribers.add(fn); return () => subscribers.delete(fn); };
    return [get, set, subscribe];
}
function computed(deps, fn) {
    const [get, set] = createSignal(fn(...deps.map(d => d())));
    deps.forEach(dep => dep[2](() => set(fn(...deps.map(d => d())))));
    return get;
}
const [count, setCount, subCount] = createSignal(0);
const doubled = computed([[count, null, subCount]], c => c * 2);
const log = [];
subCount(v => log.push("count:" + v));
setCount(5);
setCount(10);
console.log(log.join(","));
console.log(doubled());
