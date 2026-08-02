// vybe-test: js/async_concurrency_patterns/async_queue_sequential
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

class AsyncQueue {
    #queue = [];
    #running = false;
    enqueue(task) {
        return new Promise((resolve, reject) => {
            this.#queue.push({ task, resolve, reject });
            this.#run();
        });
    }
    async #run() {
        if (this.#running) return;
        this.#running = true;
        while (this.#queue.length) {
            const { task, resolve, reject } = this.#queue.shift();
            try { resolve(await task()); } catch(e) { reject(e); }
        }
        this.#running = false;
    }
}
async function main() {
    const q = new AsyncQueue();
    const order = [];
    const results = await Promise.all([
        q.enqueue(async () => { order.push(1); return "a"; }),
        q.enqueue(async () => { order.push(2); return "b"; }),
        q.enqueue(async () => { order.push(3); return "c"; }),
    ]);
    console.log(results.join(","));
    console.log(order.join(","));
}
main();
