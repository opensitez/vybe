// vybe-test: js/reactive_patterns/simple_observable
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

class Observable {
    constructor(subscribe) { this._subscribe = subscribe; }
    subscribe(observer) { return this._subscribe(observer); }
    map(fn) {
        return new Observable(obs => this.subscribe({
            next: v => obs.next(fn(v)),
            error: e => obs.error(e),
            complete: () => obs.complete()
        }));
    }
    filter(fn) {
        return new Observable(obs => this.subscribe({
            next: v => fn(v) && obs.next(v),
            error: e => obs.error(e),
            complete: () => obs.complete()
        }));
    }
    static of(...values) {
        return new Observable(obs => {
            values.forEach(v => obs.next(v));
            obs.complete();
        });
    }
}
const results = [];
Observable.of(1,2,3,4,5)
    .filter(x => x % 2 === 0)
    .map(x => x * 10)
    .subscribe({ next: v => results.push(v), error: ()=>{}, complete: ()=>{} });
console.log(results.join(","));
