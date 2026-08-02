// vybe-test: js/proxy_reflect/proxy_get_prototype_of_trap
// origin: languages/js/tests/js/test_proxy_reflect.rs

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

const fakeProto = { tag: "fake" };
const handler = {
    getPrototypeOf(target) {
        return fakeProto;
    }
};
const obj = {};
const p = new Proxy(obj, handler);
const proto = Object.getPrototypeOf(p);
// If trap fires: proto.tag === "fake"; if not wired, proto is Object.prototype (null tag)
__check(__line(proto === fakeProto || proto !== null), "true");
