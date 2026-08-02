// vybe-test: js/async_utility_patterns/sequential_async_queue
// origin: languages/js/tests/js/test_async_utility_patterns.rs

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

class AsyncQueue {
    #queue = [];
    #running = false;
    add(fn) {
        return new Promise((resolve, reject) => {
            this.#queue.push({ fn, resolve, reject });
            this.#run();
        });
    }
    async #run() {
        if (this.#running) return;
        this.#running = true;
        while (this.#queue.length > 0) {
            const { fn, resolve, reject } = this.#queue.shift();
            try { resolve(await fn()); }
            catch (e) { reject(e); }
        }
        this.#running = false;
    }
}
async function main() {
    const q = new AsyncQueue();
    const log = [];
    await Promise.all([
        q.add(async () => { log.push(1); return 1; }),
        q.add(async () => { log.push(2); return 2; }),
        q.add(async () => { log.push(3); return 3; }),
    ]);
    console.log(log.join(","));
}
main();
