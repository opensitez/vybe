function check(condition, message) {
    if (!condition) throw new Error(message);
}
const child = {identity: 1};
const target = {first: 0, second: 0};
const result = Object.assign(target, {first: child, second: 2}, {first: 3, added: 4});
check(result === target, 'target identity');
check(target.first === 3 && target.second === 2 && target.added === 4, 'all sources visited');
check(Object.keys(target).join(',') === 'first,second,added', 'existing order and new-key order');
const alias = {first: child, second: 2};
Object.assign(alias, alias);
check(alias.first === child && alias.second === 2, 'self-source identity');
const sealed = {first: 0, second: 0};
Object.preventExtensions(sealed);
Object.assign(sealed, {first: 7, second: child});
check(sealed.first === 7 && sealed.second === child, 'non-extensible existing writes');
const hidden = {};
Object.defineProperty(hidden, 'first', {value: 0, writable: true, enumerable: false});
Object.assign(hidden, {first: child});
check(hidden.first === child && Object.keys(hidden).length === 0, 'target enumerability preserved');
const mixed = {first: 0};
Object.defineProperty(mixed, 'second', {value: 1, writable: false, enumerable: true});
let thrown = false;
try { Object.assign(mixed, {first: 7, second: 8}); }
catch (error) { thrown = error instanceof TypeError; }
check(thrown && mixed.first === 7 && mixed.second === 1, 'readonly fallback preserves partial writes');
console.log('assignment batch gates passed');
