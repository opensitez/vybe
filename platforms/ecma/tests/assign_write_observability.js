// Semantic gates for an Object.assign/property-write optimization.
// Run with vybex; no Node-specific APIs or timing assertions.
function check(condition, message) {
    if (!condition) throw new Error(message);
}

function getterMutationAndInheritedSetter() {
    const events = [];
    const prototype = {};
    Object.defineProperty(prototype, "first", {
        set: function(value) { events.push("set:" + value); }
    });
    const target = Object.create(prototype);
    const source = {};
    Object.defineProperty(source, "first", {
        enumerable: true,
        get: function() {
            events.push("get");
            source.later = 29;
            source.added = 41;
            return 17;
        }
    });
    source.later = 3;
    check(Object.assign(target, source) === target, "target identity");
    check(events.join(",") === "get,set:17", "Get then inherited Set ordering");
    check(target.later === 29, "later value must be read after earlier getter");
    check(!Object.prototype.hasOwnProperty.call(target, "added"), "new key is outside initial key snapshot");
    check(!Object.prototype.hasOwnProperty.call(target, "first"), "inherited setter must not become own data property");
}

function enumerableSymbolsAndOwnKeys() {
    const inherited = { inherited: 7 };
    const source = Object.create(inherited);
    const symbol = Symbol("visible");
    const hiddenSymbol = Symbol("hidden");
    source.normal = 11;
    source[symbol] = 13;
    Object.defineProperty(source, "hidden", { value: 17, enumerable: false });
    Object.defineProperty(source, hiddenSymbol, { value: 19, enumerable: false });
    const target = Object.assign({}, source);
    check(target.normal === 11 && target[symbol] === 13, "enumerable string and symbol keys");
    check(!Object.prototype.hasOwnProperty.call(target, "inherited"), "inherited source keys excluded");
    check(!Object.prototype.hasOwnProperty.call(target, "hidden"), "non-enumerable string excluded");
    check(!Object.prototype.hasOwnProperty.call(target, hiddenSymbol), "non-enumerable symbol excluded");
}

function sourceThrowKeepsPartialWrites() {
    const sentinel = {};
    const source = { before: 5 };
    Object.defineProperty(source, "failure", {
        enumerable: true,
        get: function() { throw sentinel; }
    });
    source.after = 9;
    const target = {};
    let caught = false;
    try { Object.assign(target, source); }
    catch (error) { caught = error === sentinel; }
    check(caught, "source getter exception identity");
    check(target.before === 5, "earlier source write survives exception");
    check(!Object.prototype.hasOwnProperty.call(target, "after"), "later keys not visited after exception");
}

function targetThrowKeepsPartialWrites() {
    const sentinel = {};
    const target = {};
    Object.defineProperty(target, "failure", {
        set: function() { throw sentinel; }
    });
    let caught = false;
    try { Object.assign(target, { before: 5, failure: 7, after: 9 }); }
    catch (error) { caught = error === sentinel; }
    check(caught, "target setter exception identity");
    check(target.before === 5, "earlier target write survives exception");
    check(!Object.prototype.hasOwnProperty.call(target, "after"), "later writes stop after setter exception");
}

function nonExtensibleTarget() {
    const target = { existing: 1 };
    Object.preventExtensions(target);
    let caught = false;
    try { Object.assign(target, { existing: 2, added: 3 }); }
    catch (error) { caught = error instanceof TypeError; }
    check(caught, "adding to non-extensible target throws TypeError");
    check(target.existing === 2, "writable existing property is updated first");
    check(!Object.prototype.hasOwnProperty.call(target, "added"), "failed extension does not create property");
}

function readOnlyTarget() {
    const target = {};
    Object.defineProperty(target, "locked", { value: 1, writable: false });
    let caught = false;
    try { Object.assign(target, { before: 2, locked: 3, after: 4 }); }
    catch (error) { caught = error instanceof TypeError; }
    check(caught, "non-writable assignment throws TypeError");
    check(target.before === 2 && target.locked === 1, "partial writes and locked value preserved");
    check(!Object.prototype.hasOwnProperty.call(target, "after"), "readonly failure stops subsequent writes");
}

function ownDataShadowsInheritedSetter() {
    const events = [];
    const prototype = {};
    Object.defineProperty(prototype, "field", {
        set: function(value) { events.push(value); }
    });
    const target = Object.create(prototype);
    Object.defineProperty(target, "field", {
        value: 1, writable: true, enumerable: true, configurable: true
    });
    target.field = 2;
    check(target.field === 2 && events.length === 0, "ordinary own data write shadows prototype setter");
    Object.assign(target, { field: 3 });
    check(target.field === 3 && events.length === 0, "assign own data write shadows prototype setter");
}

function targetSetterMutatesLaterSourceValue() {
    const source = { first: 17, later: 3 };
    const target = {};
    Object.defineProperty(target, "first", {
        set: function() { source.later = 29; }
    });
    Object.assign(target, source);
    check(target.later === 29, "later source read must follow earlier target setter");
}

const checks = [getterMutationAndInheritedSetter, enumerableSymbolsAndOwnKeys,
    sourceThrowKeepsPartialWrites, targetThrowKeepsPartialWrites,
    nonExtensibleTarget, readOnlyTarget, ownDataShadowsInheritedSetter,
    targetSetterMutatesLaterSourceValue];
let failures = 0;
let index = 0;
while (index < checks.length) {
    try {
        checks[index]();
        console.log("assign gate " + index + " passed");
    } catch (error) {
        failures++;
        console.log("assign gate " + index + " failed: " + error);
    }
    index++;
}
check(failures === 0, "Object.assign semantic gates failed: " + failures);
console.log("Object.assign write observability checks passed");
