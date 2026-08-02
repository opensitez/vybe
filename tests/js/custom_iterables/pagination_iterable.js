// vybe-test: js/custom_iterables/pagination_iterable
// origin: languages/js/tests/js/test_custom_iterables.rs

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

class Paginator {
    constructor(data, pageSize) { this.data = data; this.pageSize = pageSize; }
    [Symbol.iterator]() {
        let offset = 0;
        const { data, pageSize } = this;
        return {
            next() {
                if (offset >= data.length) return { done: true };
                const page = data.slice(offset, offset + pageSize);
                offset += pageSize;
                return { value: page, done: false };
            }
        };
    }
}
const pages = [...new Paginator([1,2,3,4,5,6,7], 3)];
__check(__line(pages.length), "3");
__check(__line(pages[0].join(",")), "1,2,3");
__check(__line(pages[1].join(",")), "4,5,6");
__check(__line(pages[2].join(",")), "7");
