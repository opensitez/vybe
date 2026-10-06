function check(condition, message) {
    if (!condition) throw new Error(message);
}
function inheritedProxySet() {
    const events = [];
    const prototype = new Proxy({}, {
        set: function(target, key, value, receiver) {
            events.push(key);
            events.seenReceiver = receiver;
            return Reflect.set(target, key, value, receiver);
        }
    });
    const destination = Object.create(prototype);
    Object.assign(destination, {field: 7});
    check(events.length === 1 && events[0] === 'field', 'prototype proxy set trap observed');
    check(events.seenReceiver === destination, 'prototype proxy receives original receiver');
    check(destination.field === 7 && Object.prototype.hasOwnProperty.call(destination, 'field'),
        'forwarded set creates receiver property');
}
function inheritedProxyReject() {
    const events = [];
    const prototype = new Proxy({}, {
        set: function(target, key) { events.push(key); return false; }
    });
    const destination = Object.create(prototype);
    let thrown = false;
    try { Object.assign(destination, {first: 1, second: 2}); }
    catch (error) { thrown = error instanceof TypeError; }
    check(thrown && events.length === 1 && events[0] === 'first', 'proxy rejection throws and stops');
    check(!Object.prototype.hasOwnProperty.call(destination, 'first'), 'rejected write creates no property');
}
const gates = [inheritedProxySet, inheritedProxyReject];
let failures = 0;
let index = 0;
while (index < gates.length) {
    try { gates[index](); console.log('prototype proxy gate ' + index + ' passed'); }
    catch (error) { failures++; console.log('prototype proxy gate ' + index + ' failed: ' + error); }
    index++;
}
check(failures === 0, 'prototype proxy assignment failures: ' + failures);
console.log('prototype proxy assignment gates passed');
