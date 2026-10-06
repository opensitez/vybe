function check(condition, message) {
    if (!condition) throw new Error(message);
}
function inheritedSetter() {
    const source = {first: 1, later: 2};
    const prototype = {};
    let receiver;
    Object.defineProperty(prototype, 'first', {
        set: function(value) { receiver = this; source.later = value + 28; }
    });
    const target = Object.create(prototype);
    Object.assign(target, source);
    check(receiver === target, 'inherited setter receives destination');
    check(target.later === 29, 'inherited setter changes later source read');
    check(!Object.prototype.hasOwnProperty.call(target, 'first'), 'setter does not create own data');
}
function inheritedReadonly() {
    const prototype = {};
    Object.defineProperty(prototype, 'locked', {value: 3, writable: false});
    const target = Object.create(prototype);
    let thrown = false;
    try { Object.assign(target, {before: 1, locked: 7, after: 9}); }
    catch (error) { thrown = error instanceof TypeError; }
    check(thrown && target.before === 1, 'inherited readonly rejects after prior writes');
    check(target.locked === 3 && !Object.prototype.hasOwnProperty.call(target, 'after'), 'readonly stops later writes');
}
function inheritedGetterOnly() {
    const prototype = {};
    Object.defineProperty(prototype, 'locked', {get: function() { return 3; }});
    const target = Object.create(prototype);
    let thrown = false;
    try { Object.assign(target, {locked: 7}); }
    catch (error) { thrown = error instanceof TypeError; }
    check(thrown && !Object.prototype.hasOwnProperty.call(target, 'locked'), 'inherited getter-only rejects write');
}
function inheritedWritableData() {
    const prototype = {};
    Object.defineProperty(prototype, 'field', {value: 1, writable: true, enumerable: false});
    const target = Object.create(prototype);
    Object.assign(target, {field: 7});
    check(target.field === 7 && prototype.field === 1, 'write creates own property without changing prototype');
    check(Object.keys(target).join(',') === 'field', 'new own property is enumerable');
}
const gates = [inheritedSetter, inheritedReadonly, inheritedGetterOnly, inheritedWritableData];
let failures = 0;
let index = 0;
while (index < gates.length) {
    try { gates[index](); console.log('new-target gate ' + index + ' passed'); }
    catch (error) { failures++; console.log('new-target gate ' + index + ' failed: ' + error); }
    index++;
}
check(failures === 0, 'new-target assignment failures: ' + failures);
console.log('new-target assignment gates passed');
