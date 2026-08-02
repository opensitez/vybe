// vybe-test: js/async_concurrency_patterns/async_semaphore
// origin: languages/js/tests/js/test_async_concurrency_patterns.rs

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

class Semaphore {
    #count;
    #queue = [];
    constructor(count) { this.#count = count; }
    async acquire() {
        if (this.#count > 0) { this.#count--; return; }
        await new Promise(resolve => this.#queue.push(resolve));
    }
    release() {
        if (this.#queue.length) { this.#queue.shift()(); }
        else { this.#count++; }
    }
}
async function main() {
    const sem = new Semaphore(2);
    const log = [];
    async function task(id) {
        await sem.acquire();
        log.push("in:" + id);
        await Promise.resolve();
        log.push("out:" + id);
        sem.release();
    }
    await Promise.all([task(1), task(2), task(3)]);
    console.log(log.includes("in:1"));
    console.log(log.includes("in:2"));
    console.log(log.includes("out:3"));
}
main();
