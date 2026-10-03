//! Shared ECMA property-key and numeric-index helpers.
//!
//! These helpers keep the hot object/array/map access paths from repeatedly
//! allocating via `format!` or open-coding the same numeric-index parsing.

use std::borrow::Cow;
use std::sync::{Arc, OnceLock};
use vybe_runtime::Value;

const SMALL_INDEX_KEYS: [&str; 32] = [
    "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16",
    "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31",
];
const SMALL_INDEX_ARC_LIMIT: usize = 256;

static ASCII_CHAR_VALUES: OnceLock<[Arc<str>; 128]> = OnceLock::new();
static SMALL_INDEX_ARCS: OnceLock<[Arc<str>; SMALL_INDEX_ARC_LIMIT]> = OnceLock::new();
static HOT_STRING_ARCS: OnceLock<HotStringArcs> = OnceLock::new();

struct HotStringArcs {
    empty: Arc<str>,
    zero: Arc<str>,
    negative_zero: Arc<str>,
    true_: Arc<str>,
    false_: Arc<str>,
    null: Arc<str>,
    undefined: Arc<str>,
    length: Arc<str>,
    prototype: Arc<str>,
    constructor: Arc<str>,
    value: Arc<str>,
    to_string: Arc<str>,
    value_of: Arc<str>,
    name: Arc<str>,
    anonymous: Arc<str>,
    native_function: Arc<str>,
    object_tag: Arc<str>,
    undefined_tag: Arc<str>,
    null_tag: Arc<str>,
    boolean_tag: Arc<str>,
    number_tag: Arc<str>,
    bigint_tag: Arc<str>,
    string_tag: Arc<str>,
    symbol_tag: Arc<str>,
    array_tag: Arc<str>,
    function_tag: Arc<str>,
    map_tag: Arc<str>,
    set_tag: Arc<str>,
    date_tag: Arc<str>,
    regexp_tag: Arc<str>,
    promise_tag: Arc<str>,
    array_buffer_tag: Arc<str>,
    typed_array_tag: Arc<str>,
    string: Arc<str>,
    function: Arc<str>,
    array: Arc<str>,
    object: Arc<str>,
    number: Arc<str>,
    boolean: Arc<str>,
    symbol: Arc<str>,
    bigint: Arc<str>,
    map: Arc<str>,
    set: Arc<str>,
    date: Arc<str>,
    regexp: Arc<str>,
    promise: Arc<str>,
    ecma_array: Arc<str>,
    ecma_map: Arc<str>,
    ecma_set: Arc<str>,
    ecma_function: Arc<str>,
    ecma_object: Arc<str>,
    ecma_number: Arc<str>,
    ecma_string: Arc<str>,
    ecma_boolean: Arc<str>,
    ecma_symbol: Arc<str>,
    ecma_bigint: Arc<str>,
    ecma_date: Arc<str>,
    ecma_regexp: Arc<str>,
    ecma_json: Arc<str>,
    ecma_reflect: Arc<str>,
    ecma_promise: Arc<str>,
    ecma_iterator: Arc<str>,
    ecma_typedarray: Arc<str>,
    ecma_arraybuffer: Arc<str>,
    ecma_sharedarraybuffer: Arc<str>,
    ecma_dataview: Arc<str>,
    ecma_atomics: Arc<str>,
    ecma_math: Arc<str>,
    ecma_error: Arc<str>,
    ecma_value: Arc<str>,
    ecma_global: Arc<str>,
    ecma_global_this: Arc<str>,
    ecma_weakmap: Arc<str>,
    ecma_weakset: Arc<str>,
    ecma_weakref: Arc<str>,
    ecma_finalizationregistry: Arc<str>,
    ecma_finalization_registry: Arc<str>,
    ecma_structured_clone: Arc<str>,
    ecma_fixedarray: Arc<str>,
    ecma_generator: Arc<str>,
    ecma_intl: Arc<str>,
    ecma_intl_collator: Arc<str>,
    ecma_intl_numberformat: Arc<str>,
    ecma_intl_datetimeformat: Arc<str>,
    ecma_intl_listformat: Arc<str>,
    ecma_intl_pluralrules: Arc<str>,
    ecma_intl_relativetimeformat: Arc<str>,
    ecma_intl_segmenter: Arc<str>,
    ecma_intl_locale: Arc<str>,
    ecma_intl_displaynames: Arc<str>,
    ecma_intl_durationformat: Arc<str>,
    ecma_intl_timezone: Arc<str>,
    ecma_int8array: Arc<str>,
    ecma_uint8array: Arc<str>,
    ecma_uint8clamped: Arc<str>,
    ecma_int16array: Arc<str>,
    ecma_uint16array: Arc<str>,
    ecma_int32array: Arc<str>,
    ecma_uint32array: Arc<str>,
    ecma_float32array: Arc<str>,
    ecma_float64array: Arc<str>,
    ecma_bigint64array: Arc<str>,
    ecma_biguint64array: Arc<str>,
    array_iterator: Arc<str>,
    iterator: Arc<str>,
    invalid_date: Arc<str>,
    aggregate_error: Arc<str>,
    error: Arc<str>,
    raw_json: Arc<str>,
    array_buffer: Arc<str>,
    shared_array_buffer: Arc<str>,
    data_view: Arc<str>,
    weak_map: Arc<str>,
    weak_set: Arc<str>,
    weak_ref: Arc<str>,
    finalization_registry: Arc<str>,
    fulfilled: Arc<str>,
    rejected: Arc<str>,
    pending: Arc<str>,
    map_lower: Arc<str>,
    ok: Arc<str>,
    timed_out: Arc<str>,
    not_equal: Arc<str>,
    ecma_proxy: Arc<str>,
    call: Arc<str>,
    apply: Arc<str>,
    revoke: Arc<str>,
    empty_regexp_source: Arc<str>,
    empty_symbol: Arc<str>,
    done: Arc<str>,
    next: Arc<str>,
    return_: Arc<str>,
    throw: Arc<str>,
    message: Arc<str>,
    cause: Arc<str>,
    stack: Arc<str>,
    source: Arc<str>,
    flags: Arc<str>,
    global: Arc<str>,
    input: Arc<str>,
    index: Arc<str>,
    groups: Arc<str>,
    indices: Arc<str>,
    has_indices: Arc<str>,
    unicode_sets: Arc<str>,
    ignore_case: Arc<str>,
    multiline: Arc<str>,
    dot_all: Arc<str>,
    unicode: Arc<str>,
    sticky: Arc<str>,
    last_index: Arc<str>,
    byte_length: Arc<str>,
    byte_offset: Arc<str>,
    buffer: Arc<str>,
    size: Arc<str>,
    then: Arc<str>,
    catch: Arc<str>,
    finally: Arc<str>,
    keys: Arc<str>,
    values: Arc<str>,
    entries: Arc<str>,
    get: Arc<str>,
    set_lower: Arc<str>,
    has: Arc<str>,
    delete: Arc<str>,
    new: Arc<str>,
    from: Arc<str>,
    of: Arc<str>,
    parse: Arc<str>,
    stringify: Arc<str>,
    is_array: Arc<str>,
    push: Arc<str>,
    pop: Arc<str>,
    shift: Arc<str>,
    unshift: Arc<str>,
    slice: Arc<str>,
    splice: Arc<str>,
    index_of: Arc<str>,
    last_index_of: Arc<str>,
    includes: Arc<str>,
    join: Arc<str>,
    filter: Arc<str>,
    reduce: Arc<str>,
    for_each: Arc<str>,
    find: Arc<str>,
    sort: Arc<str>,
    reverse: Arc<str>,
    flat: Arc<str>,
    flat_map: Arc<str>,
    fill: Arc<str>,
    copy_within: Arc<str>,
    at: Arc<str>,
    concat: Arc<str>,
    every: Arc<str>,
    some: Arc<str>,
    writable: Arc<str>,
    enumerable: Arc<str>,
    configurable: Arc<str>,
    default: Arc<str>,
    has_own_property: Arc<str>,
    property_is_enumerable: Arc<str>,
    is_prototype_of: Arc<str>,
    to_locale_string: Arc<str>,
    typed_array: Arc<str>,
    module: Arc<str>,
    undefined_upper: Arc<str>,
    null_upper: Arc<str>,
    even: Arc<str>,
    odd: Arc<str>,
    proto_marker: Arc<str>,
    type_marker: Arc<str>,
    keys_marker: Arc<str>,
    primitive_marker: Arc<str>,
    nonenum_marker: Arc<str>,
}

#[inline]
fn hot_string_arcs() -> &'static HotStringArcs {
    HOT_STRING_ARCS.get_or_init(|| HotStringArcs {
        empty: Arc::from(""),
        zero: Arc::from("0"),
        negative_zero: Arc::from("-0"),
        true_: Arc::from("true"),
        false_: Arc::from("false"),
        null: Arc::from("null"),
        undefined: Arc::from("undefined"),
        length: Arc::from("length"),
        prototype: Arc::from("prototype"),
        constructor: Arc::from("constructor"),
        value: Arc::from("value"),
        name: Arc::from("name"),
        to_string: Arc::from("toString"),
        value_of: Arc::from("valueOf"),
        anonymous: Arc::from("anonymous"),
        native_function: Arc::from("function () { [native code] }"),
        object_tag: Arc::from("[object Object]"),
        undefined_tag: Arc::from("[object Undefined]"),
        null_tag: Arc::from("[object Null]"),
        boolean_tag: Arc::from("[object Boolean]"),
        number_tag: Arc::from("[object Number]"),
        bigint_tag: Arc::from("[object BigInt]"),
        string_tag: Arc::from("[object String]"),
        symbol_tag: Arc::from("[object Symbol]"),
        array_tag: Arc::from("[object Array]"),
        function_tag: Arc::from("[object Function]"),
        map_tag: Arc::from("[object Map]"),
        set_tag: Arc::from("[object Set]"),
        date_tag: Arc::from("[object Date]"),
        regexp_tag: Arc::from("[object RegExp]"),
        promise_tag: Arc::from("[object Promise]"),
        array_buffer_tag: Arc::from("[object ArrayBuffer]"),
        typed_array_tag: Arc::from("[object TypedArray]"),
        string: Arc::from("String"),
        function: Arc::from("Function"),
        array: Arc::from("Array"),
        object: Arc::from("Object"),
        number: Arc::from("Number"),
        boolean: Arc::from("Boolean"),
        symbol: Arc::from("Symbol"),
        bigint: Arc::from("BigInt"),
        map: Arc::from("Map"),
        set: Arc::from("Set"),
        date: Arc::from("Date"),
        regexp: Arc::from("RegExp"),
        promise: Arc::from("Promise"),
        ecma_array: Arc::from("ecma:array"),
        ecma_map: Arc::from("ecma:map"),
        ecma_set: Arc::from("ecma:set"),
        ecma_function: Arc::from("ecma:function"),
        ecma_object: Arc::from("ecma:object"),
        ecma_number: Arc::from("ecma:number"),
        ecma_string: Arc::from("ecma:string"),
        ecma_boolean: Arc::from("ecma:boolean"),
        ecma_symbol: Arc::from("ecma:symbol"),
        ecma_bigint: Arc::from("ecma:bigint"),
        ecma_date: Arc::from("ecma:date"),
        ecma_regexp: Arc::from("ecma:regexp"),
        ecma_json: Arc::from("ecma:json"),
        ecma_reflect: Arc::from("ecma:reflect"),
        ecma_promise: Arc::from("ecma:promise"),
        ecma_iterator: Arc::from("ecma:iterator"),
        ecma_typedarray: Arc::from("ecma:typedarray"),
        ecma_arraybuffer: Arc::from("ecma:arraybuffer"),
        ecma_sharedarraybuffer: Arc::from("ecma:sharedarraybuffer"),
        ecma_dataview: Arc::from("ecma:dataview"),
        ecma_atomics: Arc::from("ecma:atomics"),
        ecma_math: Arc::from("ecma:math"),
        ecma_error: Arc::from("ecma:error"),
        ecma_value: Arc::from("ecma:value"),
        ecma_global: Arc::from("ecma:global"),
        ecma_global_this: Arc::from("ecma:globalThis"),
        ecma_weakmap: Arc::from("ecma:weakmap"),
        ecma_weakset: Arc::from("ecma:weakset"),
        ecma_weakref: Arc::from("ecma:weakref"),
        ecma_finalizationregistry: Arc::from("ecma:finalizationregistry"),
        ecma_finalization_registry: Arc::from("ecma:finalization-registry"),
        ecma_structured_clone: Arc::from("ecma:structured-clone"),
        ecma_fixedarray: Arc::from("ecma:fixedarray"),
        ecma_generator: Arc::from("ecma:generator"),
        ecma_intl: Arc::from("ecma:intl"),
        ecma_intl_collator: Arc::from("ecma:intl/collator"),
        ecma_intl_numberformat: Arc::from("ecma:intl/numberformat"),
        ecma_intl_datetimeformat: Arc::from("ecma:intl/datetimeformat"),
        ecma_intl_listformat: Arc::from("ecma:intl/listformat"),
        ecma_intl_pluralrules: Arc::from("ecma:intl/pluralrules"),
        ecma_intl_relativetimeformat: Arc::from("ecma:intl/relativetimeformat"),
        ecma_intl_segmenter: Arc::from("ecma:intl/segmenter"),
        ecma_intl_locale: Arc::from("ecma:intl/locale"),
        ecma_intl_displaynames: Arc::from("ecma:intl/displaynames"),
        ecma_intl_durationformat: Arc::from("ecma:intl/durationformat"),
        ecma_intl_timezone: Arc::from("ecma:intl/timezone"),
        ecma_int8array: Arc::from("ecma:int8array"),
        ecma_uint8array: Arc::from("ecma:uint8array"),
        ecma_uint8clamped: Arc::from("ecma:uint8clamped"),
        ecma_int16array: Arc::from("ecma:int16array"),
        ecma_uint16array: Arc::from("ecma:uint16array"),
        ecma_int32array: Arc::from("ecma:int32array"),
        ecma_uint32array: Arc::from("ecma:uint32array"),
        ecma_float32array: Arc::from("ecma:float32array"),
        ecma_float64array: Arc::from("ecma:float64array"),
        ecma_bigint64array: Arc::from("ecma:bigint64array"),
        ecma_biguint64array: Arc::from("ecma:biguint64array"),
        array_iterator: Arc::from("ArrayIterator"),
        iterator: Arc::from("Iterator"),
        invalid_date: Arc::from("Invalid Date"),
        aggregate_error: Arc::from("AggregateError"),
        error: Arc::from("Error"),
        raw_json: Arc::from("RawJSON"),
        array_buffer: Arc::from("ArrayBuffer"),
        shared_array_buffer: Arc::from("SharedArrayBuffer"),
        data_view: Arc::from("DataView"),
        weak_map: Arc::from("WeakMap"),
        weak_set: Arc::from("WeakSet"),
        weak_ref: Arc::from("WeakRef"),
        finalization_registry: Arc::from("FinalizationRegistry"),
        fulfilled: Arc::from("fulfilled"),
        rejected: Arc::from("rejected"),
        pending: Arc::from("pending"),
        map_lower: Arc::from("map"),
        ok: Arc::from("ok"),
        timed_out: Arc::from("timed-out"),
        not_equal: Arc::from("not-equal"),
        ecma_proxy: Arc::from("ecma:proxy"),
        call: Arc::from("call"),
        apply: Arc::from("apply"),
        revoke: Arc::from("revoke"),
        empty_regexp_source: Arc::from("(?:)"),
        empty_symbol: Arc::from("Symbol()"),
        done: Arc::from("done"),
        next: Arc::from("next"),
        return_: Arc::from("return"),
        throw: Arc::from("throw"),
        message: Arc::from("message"),
        cause: Arc::from("cause"),
        stack: Arc::from("stack"),
        source: Arc::from("source"),
        flags: Arc::from("flags"),
        global: Arc::from("global"),
        input: Arc::from("input"),
        index: Arc::from("index"),
        groups: Arc::from("groups"),
        indices: Arc::from("indices"),
        has_indices: Arc::from("hasIndices"),
        unicode_sets: Arc::from("unicodeSets"),
        ignore_case: Arc::from("ignoreCase"),
        multiline: Arc::from("multiline"),
        dot_all: Arc::from("dotAll"),
        unicode: Arc::from("unicode"),
        sticky: Arc::from("sticky"),
        last_index: Arc::from("lastIndex"),
        byte_length: Arc::from("byteLength"),
        byte_offset: Arc::from("byteOffset"),
        buffer: Arc::from("buffer"),
        size: Arc::from("size"),
        then: Arc::from("then"),
        catch: Arc::from("catch"),
        finally: Arc::from("finally"),
        keys: Arc::from("keys"),
        values: Arc::from("values"),
        entries: Arc::from("entries"),
        get: Arc::from("get"),
        set_lower: Arc::from("set"),
        has: Arc::from("has"),
        delete: Arc::from("delete"),
        new: Arc::from("new"),
        from: Arc::from("from"),
        of: Arc::from("of"),
        parse: Arc::from("parse"),
        stringify: Arc::from("stringify"),
        is_array: Arc::from("isArray"),
        push: Arc::from("push"),
        pop: Arc::from("pop"),
        shift: Arc::from("shift"),
        unshift: Arc::from("unshift"),
        slice: Arc::from("slice"),
        splice: Arc::from("splice"),
        index_of: Arc::from("indexOf"),
        last_index_of: Arc::from("lastIndexOf"),
        includes: Arc::from("includes"),
        join: Arc::from("join"),
        filter: Arc::from("filter"),
        reduce: Arc::from("reduce"),
        for_each: Arc::from("forEach"),
        find: Arc::from("find"),
        sort: Arc::from("sort"),
        reverse: Arc::from("reverse"),
        flat: Arc::from("flat"),
        flat_map: Arc::from("flatMap"),
        fill: Arc::from("fill"),
        copy_within: Arc::from("copyWithin"),
        at: Arc::from("at"),
        concat: Arc::from("concat"),
        every: Arc::from("every"),
        some: Arc::from("some"),
        writable: Arc::from("writable"),
        enumerable: Arc::from("enumerable"),
        configurable: Arc::from("configurable"),
        default: Arc::from("default"),
        has_own_property: Arc::from("hasOwnProperty"),
        property_is_enumerable: Arc::from("propertyIsEnumerable"),
        is_prototype_of: Arc::from("isPrototypeOf"),
        to_locale_string: Arc::from("toLocaleString"),
        typed_array: Arc::from("TypedArray"),
        module: Arc::from("Module"),
        undefined_upper: Arc::from("Undefined"),
        null_upper: Arc::from("Null"),
        even: Arc::from("even"),
        odd: Arc::from("odd"),
        proto_marker: Arc::from("__proto__"),
        type_marker: Arc::from("__type"),
        keys_marker: Arc::from("__keys"),
        primitive_marker: Arc::from("__primitive"),
        nonenum_marker: Arc::from("__nonenum"),
    })
}

#[inline]
pub fn hot_string_arc(text: &str) -> Option<Arc<str>> {
    let strings = hot_string_arcs();
    Some(match text {
        "" => Arc::clone(&strings.empty),
        "0" => Arc::clone(&strings.zero),
        "-0" => Arc::clone(&strings.negative_zero),
        "true" => Arc::clone(&strings.true_),
        "false" => Arc::clone(&strings.false_),
        "null" => Arc::clone(&strings.null),
        "undefined" => Arc::clone(&strings.undefined),
        "length" => Arc::clone(&strings.length),
        "prototype" => Arc::clone(&strings.prototype),
        "constructor" => Arc::clone(&strings.constructor),
        "value" => Arc::clone(&strings.value),
        "name" => Arc::clone(&strings.name),
        "toString" => Arc::clone(&strings.to_string),
        "valueOf" => Arc::clone(&strings.value_of),
        "anonymous" => Arc::clone(&strings.anonymous),
        "function () { [native code] }" => Arc::clone(&strings.native_function),
        "[object Object]" => Arc::clone(&strings.object_tag),
        "[object Undefined]" => Arc::clone(&strings.undefined_tag),
        "[object Null]" => Arc::clone(&strings.null_tag),
        "[object Boolean]" => Arc::clone(&strings.boolean_tag),
        "[object Number]" => Arc::clone(&strings.number_tag),
        "[object BigInt]" => Arc::clone(&strings.bigint_tag),
        "[object String]" => Arc::clone(&strings.string_tag),
        "[object Symbol]" => Arc::clone(&strings.symbol_tag),
        "[object Array]" => Arc::clone(&strings.array_tag),
        "[object Function]" => Arc::clone(&strings.function_tag),
        "[object Map]" => Arc::clone(&strings.map_tag),
        "[object Set]" => Arc::clone(&strings.set_tag),
        "[object Date]" => Arc::clone(&strings.date_tag),
        "[object RegExp]" => Arc::clone(&strings.regexp_tag),
        "[object Promise]" => Arc::clone(&strings.promise_tag),
        "[object ArrayBuffer]" => Arc::clone(&strings.array_buffer_tag),
        "[object TypedArray]" => Arc::clone(&strings.typed_array_tag),
        "String" => Arc::clone(&strings.string),
        "Function" => Arc::clone(&strings.function),
        "Array" => Arc::clone(&strings.array),
        "Object" => Arc::clone(&strings.object),
        "Number" => Arc::clone(&strings.number),
        "Boolean" => Arc::clone(&strings.boolean),
        "Symbol" => Arc::clone(&strings.symbol),
        "BigInt" => Arc::clone(&strings.bigint),
        "Map" => Arc::clone(&strings.map),
        "Set" => Arc::clone(&strings.set),
        "Date" => Arc::clone(&strings.date),
        "RegExp" => Arc::clone(&strings.regexp),
        "Promise" => Arc::clone(&strings.promise),
        "ecma:array" => Arc::clone(&strings.ecma_array),
        "ecma:map" => Arc::clone(&strings.ecma_map),
        "ecma:set" => Arc::clone(&strings.ecma_set),
        "ecma:function" => Arc::clone(&strings.ecma_function),
        "ecma:object" => Arc::clone(&strings.ecma_object),
        "ecma:number" => Arc::clone(&strings.ecma_number),
        "ecma:string" => Arc::clone(&strings.ecma_string),
        "ecma:boolean" => Arc::clone(&strings.ecma_boolean),
        "ecma:symbol" => Arc::clone(&strings.ecma_symbol),
        "ecma:bigint" => Arc::clone(&strings.ecma_bigint),
        "ecma:date" => Arc::clone(&strings.ecma_date),
        "ecma:regexp" => Arc::clone(&strings.ecma_regexp),
        "ecma:json" => Arc::clone(&strings.ecma_json),
        "ecma:reflect" => Arc::clone(&strings.ecma_reflect),
        "ecma:promise" => Arc::clone(&strings.ecma_promise),
        "ecma:iterator" => Arc::clone(&strings.ecma_iterator),
        "ecma:typedarray" => Arc::clone(&strings.ecma_typedarray),
        "ecma:arraybuffer" => Arc::clone(&strings.ecma_arraybuffer),
        "ecma:sharedarraybuffer" => Arc::clone(&strings.ecma_sharedarraybuffer),
        "ecma:dataview" => Arc::clone(&strings.ecma_dataview),
        "ecma:atomics" => Arc::clone(&strings.ecma_atomics),
        "ecma:math" => Arc::clone(&strings.ecma_math),
        "ecma:error" => Arc::clone(&strings.ecma_error),
        "ecma:value" => Arc::clone(&strings.ecma_value),
        "ecma:global" => Arc::clone(&strings.ecma_global),
        "ecma:globalThis" => Arc::clone(&strings.ecma_global_this),
        "ecma:weakmap" => Arc::clone(&strings.ecma_weakmap),
        "ecma:weakset" => Arc::clone(&strings.ecma_weakset),
        "ecma:weakref" => Arc::clone(&strings.ecma_weakref),
        "ecma:finalizationregistry" => Arc::clone(&strings.ecma_finalizationregistry),
        "ecma:finalization-registry" => Arc::clone(&strings.ecma_finalization_registry),
        "ecma:structured-clone" => Arc::clone(&strings.ecma_structured_clone),
        "ecma:fixedarray" => Arc::clone(&strings.ecma_fixedarray),
        "ecma:generator" => Arc::clone(&strings.ecma_generator),
        "ecma:intl" => Arc::clone(&strings.ecma_intl),
        "ecma:intl/collator" => Arc::clone(&strings.ecma_intl_collator),
        "ecma:intl/numberformat" => Arc::clone(&strings.ecma_intl_numberformat),
        "ecma:intl/datetimeformat" => Arc::clone(&strings.ecma_intl_datetimeformat),
        "ecma:intl/listformat" => Arc::clone(&strings.ecma_intl_listformat),
        "ecma:intl/pluralrules" => Arc::clone(&strings.ecma_intl_pluralrules),
        "ecma:intl/relativetimeformat" => Arc::clone(&strings.ecma_intl_relativetimeformat),
        "ecma:intl/segmenter" => Arc::clone(&strings.ecma_intl_segmenter),
        "ecma:intl/locale" => Arc::clone(&strings.ecma_intl_locale),
        "ecma:intl/displaynames" => Arc::clone(&strings.ecma_intl_displaynames),
        "ecma:intl/durationformat" => Arc::clone(&strings.ecma_intl_durationformat),
        "ecma:intl/timezone" => Arc::clone(&strings.ecma_intl_timezone),
        "ecma:int8array" => Arc::clone(&strings.ecma_int8array),
        "ecma:uint8array" => Arc::clone(&strings.ecma_uint8array),
        "ecma:uint8clamped" => Arc::clone(&strings.ecma_uint8clamped),
        "ecma:int16array" => Arc::clone(&strings.ecma_int16array),
        "ecma:uint16array" => Arc::clone(&strings.ecma_uint16array),
        "ecma:int32array" => Arc::clone(&strings.ecma_int32array),
        "ecma:uint32array" => Arc::clone(&strings.ecma_uint32array),
        "ecma:float32array" => Arc::clone(&strings.ecma_float32array),
        "ecma:float64array" => Arc::clone(&strings.ecma_float64array),
        "ecma:bigint64array" => Arc::clone(&strings.ecma_bigint64array),
        "ecma:biguint64array" => Arc::clone(&strings.ecma_biguint64array),
        "ArrayIterator" => Arc::clone(&strings.array_iterator),
        "Iterator" => Arc::clone(&strings.iterator),
        "Invalid Date" => Arc::clone(&strings.invalid_date),
        "AggregateError" => Arc::clone(&strings.aggregate_error),
        "Error" => Arc::clone(&strings.error),
        "RawJSON" => Arc::clone(&strings.raw_json),
        "ArrayBuffer" => Arc::clone(&strings.array_buffer),
        "SharedArrayBuffer" => Arc::clone(&strings.shared_array_buffer),
        "DataView" => Arc::clone(&strings.data_view),
        "WeakMap" => Arc::clone(&strings.weak_map),
        "WeakSet" => Arc::clone(&strings.weak_set),
        "WeakRef" => Arc::clone(&strings.weak_ref),
        "FinalizationRegistry" => Arc::clone(&strings.finalization_registry),
        "fulfilled" => Arc::clone(&strings.fulfilled),
        "rejected" => Arc::clone(&strings.rejected),
        "pending" => Arc::clone(&strings.pending),
        "map" => Arc::clone(&strings.map_lower),
        "ok" => Arc::clone(&strings.ok),
        "timed-out" => Arc::clone(&strings.timed_out),
        "not-equal" => Arc::clone(&strings.not_equal),
        "ecma:proxy" => Arc::clone(&strings.ecma_proxy),
        "call" => Arc::clone(&strings.call),
        "apply" => Arc::clone(&strings.apply),
        "revoke" => Arc::clone(&strings.revoke),
        "(?:)" => Arc::clone(&strings.empty_regexp_source),
        "Symbol()" => Arc::clone(&strings.empty_symbol),
        "done" => Arc::clone(&strings.done),
        "next" => Arc::clone(&strings.next),
        "return" => Arc::clone(&strings.return_),
        "throw" => Arc::clone(&strings.throw),
        "message" => Arc::clone(&strings.message),
        "cause" => Arc::clone(&strings.cause),
        "stack" => Arc::clone(&strings.stack),
        "source" => Arc::clone(&strings.source),
        "flags" => Arc::clone(&strings.flags),
        "global" => Arc::clone(&strings.global),
        "input" => Arc::clone(&strings.input),
        "index" => Arc::clone(&strings.index),
        "groups" => Arc::clone(&strings.groups),
        "indices" => Arc::clone(&strings.indices),
        "hasIndices" => Arc::clone(&strings.has_indices),
        "unicodeSets" => Arc::clone(&strings.unicode_sets),
        "ignoreCase" => Arc::clone(&strings.ignore_case),
        "multiline" => Arc::clone(&strings.multiline),
        "dotAll" => Arc::clone(&strings.dot_all),
        "unicode" => Arc::clone(&strings.unicode),
        "sticky" => Arc::clone(&strings.sticky),
        "lastIndex" => Arc::clone(&strings.last_index),
        "byteLength" => Arc::clone(&strings.byte_length),
        "byteOffset" => Arc::clone(&strings.byte_offset),
        "buffer" => Arc::clone(&strings.buffer),
        "size" => Arc::clone(&strings.size),
        "then" => Arc::clone(&strings.then),
        "catch" => Arc::clone(&strings.catch),
        "finally" => Arc::clone(&strings.finally),
        "keys" => Arc::clone(&strings.keys),
        "values" => Arc::clone(&strings.values),
        "entries" => Arc::clone(&strings.entries),
        "get" => Arc::clone(&strings.get),
        "set" => Arc::clone(&strings.set_lower),
        "has" => Arc::clone(&strings.has),
        "delete" => Arc::clone(&strings.delete),
        "new" => Arc::clone(&strings.new),
        "from" => Arc::clone(&strings.from),
        "of" => Arc::clone(&strings.of),
        "parse" => Arc::clone(&strings.parse),
        "stringify" => Arc::clone(&strings.stringify),
        "isArray" => Arc::clone(&strings.is_array),
        "push" => Arc::clone(&strings.push),
        "pop" => Arc::clone(&strings.pop),
        "shift" => Arc::clone(&strings.shift),
        "unshift" => Arc::clone(&strings.unshift),
        "slice" => Arc::clone(&strings.slice),
        "splice" => Arc::clone(&strings.splice),
        "indexOf" => Arc::clone(&strings.index_of),
        "lastIndexOf" => Arc::clone(&strings.last_index_of),
        "includes" => Arc::clone(&strings.includes),
        "join" => Arc::clone(&strings.join),
        "filter" => Arc::clone(&strings.filter),
        "reduce" => Arc::clone(&strings.reduce),
        "forEach" => Arc::clone(&strings.for_each),
        "find" => Arc::clone(&strings.find),
        "sort" => Arc::clone(&strings.sort),
        "reverse" => Arc::clone(&strings.reverse),
        "flat" => Arc::clone(&strings.flat),
        "flatMap" => Arc::clone(&strings.flat_map),
        "fill" => Arc::clone(&strings.fill),
        "copyWithin" => Arc::clone(&strings.copy_within),
        "at" => Arc::clone(&strings.at),
        "concat" => Arc::clone(&strings.concat),
        "every" => Arc::clone(&strings.every),
        "some" => Arc::clone(&strings.some),
        "writable" => Arc::clone(&strings.writable),
        "enumerable" => Arc::clone(&strings.enumerable),
        "configurable" => Arc::clone(&strings.configurable),
        "default" => Arc::clone(&strings.default),
        "hasOwnProperty" => Arc::clone(&strings.has_own_property),
        "propertyIsEnumerable" => Arc::clone(&strings.property_is_enumerable),
        "isPrototypeOf" => Arc::clone(&strings.is_prototype_of),
        "toLocaleString" => Arc::clone(&strings.to_locale_string),
        "TypedArray" => Arc::clone(&strings.typed_array),
        "Module" => Arc::clone(&strings.module),
        "Undefined" => Arc::clone(&strings.undefined_upper),
        "Null" => Arc::clone(&strings.null_upper),
        "even" => Arc::clone(&strings.even),
        "odd" => Arc::clone(&strings.odd),
        "__proto__" => Arc::clone(&strings.proto_marker),
        "__type" => Arc::clone(&strings.type_marker),
        "__keys" => Arc::clone(&strings.keys_marker),
        "__primitive" => Arc::clone(&strings.primitive_marker),
        "__nonenum" => Arc::clone(&strings.nonenum_marker),
        _ => return None,
    })
}

#[inline]
fn small_index_text_index(text: &str) -> Option<usize> {
    let bytes = text.as_bytes();
    let index = match bytes {
        [b'0'] => 0,
        [b'1'..=b'9'] => (bytes[0] - b'0') as usize,
        [b'1'..=b'9', b'0'..=b'9'] => {
            ((bytes[0] - b'0') as usize) * 10 + (bytes[1] - b'0') as usize
        }
        [b'1', b'0'..=b'9', b'0'..=b'9'] => {
            100 + ((bytes[1] - b'0') as usize) * 10 + (bytes[2] - b'0') as usize
        }
        [b'2', b'0'..=b'4', b'0'..=b'9'] => {
            200 + ((bytes[1] - b'0') as usize) * 10 + (bytes[2] - b'0') as usize
        }
        [b'2', b'5', b'0'..=b'5'] => 250 + (bytes[2] - b'0') as usize,
        _ => return None,
    };
    Some(index)
}

#[inline]
pub fn small_index_text_arc(text: &str) -> Option<Arc<str>> {
    small_index_text_index(text).and_then(small_index_arc)
}

#[inline]
pub fn string_arc(text: &str) -> Arc<str> {
    hot_string_arc(text)
        .or_else(|| small_index_text_arc(text))
        .unwrap_or_else(|| Arc::from(text))
}

#[inline]
pub fn owned_string_arc(text: String) -> Arc<str> {
    if let Some(cached) = hot_string_arc(&text).or_else(|| small_index_text_arc(&text)) {
        cached
    } else {
        Arc::from(text)
    }
}

#[inline]
pub fn string_value(text: &str) -> Value {
    Value::String(string_arc(text))
}

#[inline]
pub fn owned_string_value(text: String) -> Value {
    Value::String(owned_string_arc(text))
}

#[inline]
pub fn char_value(ch: char) -> Value {
    if ch.is_ascii() {
        let values = ASCII_CHAR_VALUES.get_or_init(|| {
            std::array::from_fn(|index| {
                let byte = index as u8;
                let text = std::str::from_utf8(std::slice::from_ref(&byte)).unwrap_or("");
                Arc::<str>::from(text)
            })
        });
        Value::String(Arc::clone(&values[ch as usize]))
    } else {
        let mut buf = [0u8; 4];
        Value::String(Arc::from(ch.encode_utf8(&mut buf)))
    }
}

#[inline]
pub fn small_index_key(index: usize) -> Option<&'static str> {
    SMALL_INDEX_KEYS.get(index).copied()
}

#[inline]
pub fn small_index_arc(index: usize) -> Option<Arc<str>> {
    if index >= SMALL_INDEX_ARC_LIMIT {
        return None;
    }
    let values = SMALL_INDEX_ARCS
        .get_or_init(|| std::array::from_fn(|index| Arc::<str>::from(index.to_string())));
    Some(Arc::clone(&values[index]))
}

#[inline]
pub fn small_index_string_value(index: usize) -> Option<Value> {
    small_index_arc(index).map(Value::String)
}

#[inline]
pub fn with_index_key<R>(index: usize, f: impl FnOnce(&str) -> R) -> R {
    if let Some(key) = small_index_key(index) {
        return f(key);
    }
    let mut buf = [0u8; 20];
    let mut value = index;
    let mut pos = buf.len();
    if value == 0 {
        pos -= 1;
        buf[pos] = b'0';
    } else {
        while value > 0 {
            pos -= 1;
            buf[pos] = b'0' + (value % 10) as u8;
            value /= 10;
        }
    }
    let key = std::str::from_utf8(&buf[pos..]).unwrap_or("");
    f(key)
}

#[inline]
pub fn prefixed_property_key(prefix: &str, key: &str) -> String {
    let mut out = String::with_capacity(prefix.len() + key.len());
    out.push_str(prefix);
    out.push_str(key);
    out
}

#[inline]
pub fn with_prefixed_property_key<R>(prefix: &str, key: &str, f: impl FnOnce(&str) -> R) -> R {
    const INLINE_KEY_BYTES: usize = 128;
    let len = prefix.len() + key.len();
    if len <= INLINE_KEY_BYTES {
        let mut buf = [0u8; INLINE_KEY_BYTES];
        buf[..prefix.len()].copy_from_slice(prefix.as_bytes());
        buf[prefix.len()..len].copy_from_slice(key.as_bytes());
        let text = std::str::from_utf8(&buf[..len]).unwrap_or("");
        f(text)
    } else {
        let owned = prefixed_property_key(prefix, key);
        f(&owned)
    }
}

#[inline]
pub fn with_wrapped_property_key<R>(
    prefix: &str,
    key: &str,
    suffix: &str,
    f: impl FnOnce(&str) -> R,
) -> R {
    const INLINE_KEY_BYTES: usize = 128;
    let len = prefix.len() + key.len() + suffix.len();
    if len <= INLINE_KEY_BYTES {
        let mut buf = [0u8; INLINE_KEY_BYTES];
        buf[..prefix.len()].copy_from_slice(prefix.as_bytes());
        buf[prefix.len()..prefix.len() + key.len()].copy_from_slice(key.as_bytes());
        buf[prefix.len() + key.len()..len].copy_from_slice(suffix.as_bytes());
        let text = std::str::from_utf8(&buf[..len]).unwrap_or("");
        f(text)
    } else {
        let mut owned = String::with_capacity(len);
        owned.push_str(prefix);
        owned.push_str(key);
        owned.push_str(suffix);
        f(&owned)
    }
}

#[inline]
pub fn getter_property_key(key: &str) -> String {
    prefixed_property_key("__get_", key)
}

#[inline]
pub fn setter_property_key(key: &str) -> String {
    prefixed_property_key("__set_", key)
}

#[inline]
pub fn with_getter_property_key<R>(key: &str, f: impl FnOnce(&str) -> R) -> R {
    with_prefixed_property_key("__get_", key, f)
}

#[inline]
pub fn with_setter_property_key<R>(key: &str, f: impl FnOnce(&str) -> R) -> R {
    with_prefixed_property_key("__set_", key, f)
}

#[inline]
pub fn property_key_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.as_ref().to_owned(),
        Value::Symbol(sym) => crate::symbol::canonical_property_key(sym),
        Value::I32(n) if *n >= 0 => with_index_key(*n as usize, str::to_owned),
        Value::I32(n) => n.to_string(),
        Value::I64(n) if *n >= 0 => match usize::try_from(*n) {
            Ok(index) => with_index_key(index, str::to_owned),
            Err(_) => n.to_string(),
        },
        Value::I64(n) => n.to_string(),
        Value::F32(n) => n.to_string(),
        Value::F64(n) => n.to_string(),
        Value::Bool(true) => "true".to_string(),
        Value::Bool(false) => "false".to_string(),
        Value::Null | Value::TypedNull(_) => "null".to_string(),
        Value::Undefined => "undefined".to_string(),
        Value::BigInt(n) => n.to_string(),
        other => format!("{}", other),
    }
}

#[inline]
pub fn value_display_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.as_ref().to_owned(),
        Value::I32(n) if *n >= 0 => with_index_key(*n as usize, str::to_owned),
        Value::I32(n) => n.to_string(),
        Value::I64(n) if *n >= 0 => match usize::try_from(*n) {
            Ok(index) => with_index_key(index, str::to_owned),
            Err(_) => n.to_string(),
        },
        Value::I64(n) => n.to_string(),
        Value::F32(n) => n.to_string(),
        Value::F64(n) => n.to_string(),
        Value::Bool(true) => "true".to_string(),
        Value::Bool(false) => "false".to_string(),
        Value::Null | Value::TypedNull(_) => "null".to_string(),
        Value::Undefined => "undefined".to_string(),
        Value::BigInt(n) => n.to_string(),
        other => format!("{}", other),
    }
}

#[inline]
pub fn value_display_cow<'a>(value: &'a Value) -> Cow<'a, str> {
    match value {
        Value::String(text) => Cow::Borrowed(text.as_ref()),
        Value::I32(n) if *n >= 0 => small_index_key(*n as usize)
            .map(Cow::Borrowed)
            .unwrap_or_else(|| Cow::Owned(n.to_string())),
        Value::I32(n) => Cow::Owned(n.to_string()),
        Value::I64(n) if *n >= 0 => usize::try_from(*n)
            .ok()
            .and_then(small_index_key)
            .map(Cow::Borrowed)
            .unwrap_or_else(|| Cow::Owned(n.to_string())),
        Value::I64(n) => Cow::Owned(n.to_string()),
        Value::F32(n) => Cow::Owned(n.to_string()),
        Value::F64(n) => Cow::Owned(n.to_string()),
        Value::Bool(true) => Cow::Borrowed("true"),
        Value::Bool(false) => Cow::Borrowed("false"),
        Value::Null | Value::TypedNull(_) => Cow::Borrowed("null"),
        Value::Undefined => Cow::Borrowed("undefined"),
        Value::BigInt(n) => Cow::Owned(n.to_string()),
        other => Cow::Owned(format!("{}", other)),
    }
}

#[inline]
pub fn concat2_arc(left: &str, right: &str) -> Arc<str> {
    let mut out = String::with_capacity(left.len() + right.len());
    out.push_str(left);
    out.push_str(right);
    Arc::<str>::from(out)
}

#[inline]
pub fn concat3_arc(left: &str, middle: &str, right: &str) -> Arc<str> {
    let mut out = String::with_capacity(left.len() + middle.len() + right.len());
    out.push_str(left);
    out.push_str(middle);
    out.push_str(right);
    Arc::<str>::from(out)
}

#[inline]
pub fn object_tag_value(tag: &str) -> Value {
    match tag {
        "Object" => string_value("[object Object]"),
        "Undefined" => string_value("[object Undefined]"),
        "Null" => string_value("[object Null]"),
        "Boolean" => string_value("[object Boolean]"),
        "Number" => string_value("[object Number]"),
        "BigInt" => string_value("[object BigInt]"),
        "String" => string_value("[object String]"),
        "Symbol" => string_value("[object Symbol]"),
        "Array" => string_value("[object Array]"),
        "Function" => string_value("[object Function]"),
        "Map" => string_value("[object Map]"),
        "Set" => string_value("[object Set]"),
        "Date" => string_value("[object Date]"),
        "RegExp" => string_value("[object RegExp]"),
        "Promise" => string_value("[object Promise]"),
        "ArrayBuffer" => string_value("[object ArrayBuffer]"),
        "TypedArray" => string_value("[object TypedArray]"),
        _ => Value::String(concat3_arc("[object ", tag, "]")),
    }
}

#[inline]
pub fn with_property_key<R>(value: &Value, f: impl FnOnce(&str) -> R) -> R {
    match value {
        Value::String(text) => f(text.as_ref()),
        Value::Symbol(sym) => {
            let key = crate::symbol::canonical_property_key(sym);
            f(&key)
        }
        Value::I32(n) if *n >= 0 => with_index_key(*n as usize, f),
        Value::I64(n) if *n >= 0 => match usize::try_from(*n) {
            Ok(index) => with_index_key(index, f),
            Err(_) => {
                let key = n.to_string();
                f(&key)
            }
        },
        Value::F32(n)
            if n.is_finite()
                && n.fract() == 0.0
                && *n >= 0.0
                && *n <= 16_777_216.0
                && !(*n == 0.0 && n.is_sign_negative()) =>
        {
            with_index_key(*n as usize, f)
        }
        Value::F64(n)
            if n.is_finite()
                && n.fract() == 0.0
                && *n >= 0.0
                && *n <= 9_007_199_254_740_991.0
                && !(*n == 0.0 && n.is_sign_negative()) =>
        {
            with_index_key(*n as usize, f)
        }
        Value::Bool(true) => f("true"),
        Value::Bool(false) => f("false"),
        Value::Null | Value::TypedNull(_) => f("null"),
        Value::Undefined => f("undefined"),
        Value::BigInt(n) => {
            let key = n.to_string();
            f(&key)
        }
        other => {
            let key = format!("{}", other);
            f(&key)
        }
    }
}

#[inline]
pub fn map_lookup_key_cow(value: &Value) -> Cow<'_, Value> {
    match value {
        Value::String(_) | Value::I32(_) | Value::I64(_) | Value::F32(_) | Value::F64(_) => {
            Cow::Borrowed(value)
        }
        _ => Cow::Owned(map_lookup_key(value)),
    }
}

#[inline]
pub fn map_lookup_key(value: &Value) -> Value {
    match value {
        Value::String(_) | Value::I32(_) | Value::I64(_) | Value::F32(_) | Value::F64(_) => {
            value.clone()
        }
        Value::Bool(true) => string_value("true"),
        Value::Bool(false) => string_value("false"),
        Value::Null | Value::TypedNull(_) => string_value("null"),
        Value::Undefined => string_value("undefined"),
        other => string_value(&property_key_string(other)),
    }
}

#[inline]
pub fn non_negative_integer_index(value: &Value) -> Option<usize> {
    match value {
        Value::I32(n) if *n >= 0 => Some(*n as usize),
        Value::I64(n) if *n >= 0 => usize::try_from(*n).ok(),
        Value::F32(n) if n.is_finite() && n.fract() == 0.0 && *n >= 0.0 => Some(*n as usize),
        Value::F64(n) if n.is_finite() && n.fract() == 0.0 && *n >= 0.0 => Some(*n as usize),
        Value::String(text) => non_negative_integer_index_key(text),
        _ => None,
    }
}

#[inline]
pub fn non_negative_integer_index_key(key: &str) -> Option<usize> {
    let bytes = key.as_bytes();
    if bytes.is_empty() {
        return None;
    }
    let mut value: usize = 0;
    for &byte in bytes {
        if !byte.is_ascii_digit() {
            return None;
        }
        value = value.checked_mul(10)?.checked_add((byte - b'0') as usize)?;
    }
    Some(value)
}

#[inline]
pub fn non_negative_u32_key(key: &str) -> Option<u32> {
    let value = non_negative_integer_index_key(key)?;
    u32::try_from(value).ok()
}

#[inline]
pub fn non_negative_i32_key(key: &str) -> Option<i32> {
    let value = non_negative_integer_index_key(key)?;
    i32::try_from(value).ok()
}

#[inline]
pub fn canonical_array_index_key(key: &str) -> Option<u32> {
    if key.is_empty() {
        None
    } else if key == "0" {
        Some(0)
    } else if key.as_bytes().first() == Some(&b'0') {
        None
    } else {
        let mut value: u64 = 0;
        for byte in key.bytes() {
            if !byte.is_ascii_digit() {
                return None;
            }
            value = value
                .saturating_mul(10)
                .saturating_add((byte - b'0') as u64);
            if value >= u32::MAX as u64 {
                return None;
            }
        }
        Some(value as u32)
    }
}

#[inline]
pub fn canonical_array_index_value(value: &Value) -> Option<u32> {
    match value {
        Value::I32(n) if *n >= 0 && (*n as u32) != u32::MAX => Some(*n as u32),
        Value::I64(n) if *n >= 0 && *n < u32::MAX as i64 => Some(*n as u32),
        Value::F32(n)
            if n.is_finite()
                && n.fract() == 0.0
                && *n >= 0.0
                && *n < u32::MAX as f32
                && !(*n == 0.0 && n.is_sign_negative()) =>
        {
            Some(*n as u32)
        }
        Value::F64(n)
            if n.is_finite()
                && n.fract() == 0.0
                && *n >= 0.0
                && *n < u32::MAX as f64
                && !(*n == 0.0 && n.is_sign_negative()) =>
        {
            Some(*n as u32)
        }
        Value::String(text) => canonical_array_index_key(text),
        _ => None,
    }
}

#[inline]
pub fn canonical_numeric_index_value(value: &Value) -> Option<f64> {
    match value {
        Value::I32(n) => Some(*n as f64),
        Value::I64(n) => Some(*n as f64),
        Value::F32(n) => Some(*n as f64),
        Value::F64(n) => Some(*n),
        Value::String(text) if text.as_ref() == "-0" => Some(-0.0),
        Value::String(text) => {
            if let Some(index) = canonical_array_index_key(text) {
                Some(index as f64)
            } else {
                let number = text.parse::<f64>().ok()?;
                if Value::F64(number).to_string() == text.as_ref() {
                    Some(number)
                } else {
                    None
                }
            }
        }
        _ => None,
    }
}
