// Separate compiler/runtime diagnostic, not an ECMA assignment fast-path gate.
// A closed-over lexical write must remain visible after the callback returns.
function closureWriteback() {
    let observed;
    const expected = {};
    const callback = function(value) { observed = value; };
    callback(expected);
    if (observed !== expected) throw new Error('closed-over lexical writeback lost');
}
closureWriteback();
console.log('direct closure writeback passed');
function declaredHandlerWriteback() {
    let observed;
    const handler = {
        set: function(target, key, value, receiver) {
            observed = receiver;
            return Reflect.set(target, key, value, receiver);
        }
    };
    const proxy = new Proxy({}, handler);
    Object.assign(proxy, {field: 7});
    if (observed !== proxy) throw new Error('declared handler lexical writeback lost');
}
declaredHandlerWriteback();
console.log('declared handler closure writeback passed');
function proxyClosureWriteback() {
    let observed;
    const proxy = new Proxy({}, {
        set: function(target, key, value, receiver) {
            observed = receiver;
            return Reflect.set(target, key, value, receiver);
        }
    });
    Object.assign(proxy, {field: 7});
    if (observed !== proxy) throw new Error('inline constructor handler lexical writeback lost');
}
proxyClosureWriteback();
console.log('proxy callback closure writeback passed');
