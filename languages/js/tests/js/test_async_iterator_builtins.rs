crate::js_cases! {
    promise_withresolvers_exposes_capabilities => {
        r#"
const result = Promise.withResolvers();
console.log(typeof result.promise);
console.log(typeof result.resolve);
console.log(typeof result.reject);
"#,
        ["object", "function", "function"]
    };

    promise_withresolvers_resolves_created_promise => {
        r#"
const { promise, resolve } = Promise.withResolvers();
resolve(42);
console.log(await promise);
"#,
        ["42"]
    };

    promise_withresolvers_rejects_created_promise => {
        r#"
const { promise, reject } = Promise.withResolvers();
reject("boom");
try {
  await promise;
} catch (error) {
  console.log(error);
}
"#,
        ["boom"]
    };

    array_fromasync_collects_async_generator_values => {
        r#"
async function* numbers() {
  yield 1;
  yield 2;
  yield 3;
}
const result = await Array.fromAsync(numbers());
console.log(result.join(","));
"#,
        ["1,2,3"]
    };

    array_fromasync_applies_mapping_function => {
        r#"
async function* numbers() {
  yield 2;
  yield 4;
}
const result = await Array.fromAsync(numbers(), value => value / 2);
console.log(result.join(","));
"#,
        ["1,2"]
    };

    array_fromasync_awaits_promises_from_sync_iterable => {
        r#"
const result = await Array.fromAsync([Promise.resolve("a"), Promise.resolve("b")]);
console.log(result.join(","));
"#,
        ["a,b"]
    };

    iterator_from_wraps_array_iterable => {
        r#"
const iterator = Iterator.from([10, 20, 30]);
console.log(iterator.next().value);
console.log(iterator.next().value);
console.log(iterator.next().value);
"#,
        ["10", "20", "30"]
    };

    iterator_from_accepts_custom_iterable_objects => {
        r#"
const source = {
  *[Symbol.iterator]() {
    yield "x";
    yield "y";
  }
};
const iterator = Iterator.from(source);
console.log(iterator.next().value + iterator.next().value);
"#,
        ["xy"]
    };

    asynciterator_from_wraps_async_iterables => {
        r#"
async function* source() {
  yield "left";
  yield "right";
}
const iterator = AsyncIterator.from(source());
console.log((await iterator.next()).value);
console.log((await iterator.next()).value);
"#,
        ["left", "right"]
    };

    asynciterator_from_wraps_sync_iterable => {
        r#"
const asyncIter = AsyncIterator.from([10, 20]);
console.log((await asyncIter.next()).value + "|" + (await asyncIter.next()).value);
"#,
        ["10|20"]
    };
}

