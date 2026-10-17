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
    // Length partitions avoid a linear chain over every cached spelling in
    // debug builds. Exact comparisons and the shared Arc identities remain.
    Some(match text.len() {
        0 => match text {
            "" => Arc::clone(&strings.empty),
            _ => return None,
        },
        1 => match text {
            "0" => Arc::clone(&strings.zero),
            _ => return None,
        },
        2 => match text {
            "-0" => Arc::clone(&strings.negative_zero),
            "ok" => Arc::clone(&strings.ok),
            "of" => Arc::clone(&strings.of),
            "at" => Arc::clone(&strings.at),
            _ => return None,
        },
        3 => match text {
            "Map" => Arc::clone(&strings.map),
            "Set" => Arc::clone(&strings.set),
            "map" => Arc::clone(&strings.map_lower),
            "get" => Arc::clone(&strings.get),
            "set" => Arc::clone(&strings.set_lower),
            "has" => Arc::clone(&strings.has),
            "new" => Arc::clone(&strings.new),
            "pop" => Arc::clone(&strings.pop),
            "odd" => Arc::clone(&strings.odd),
            _ => return None,
        },
        4 => match text {
            "true" => Arc::clone(&strings.true_),
            "null" => Arc::clone(&strings.null),
            "name" => Arc::clone(&strings.name),
            "Date" => Arc::clone(&strings.date),
            "call" => Arc::clone(&strings.call),
            "(?:)" => Arc::clone(&strings.empty_regexp_source),
            "done" => Arc::clone(&strings.done),
            "next" => Arc::clone(&strings.next),
            "size" => Arc::clone(&strings.size),
            "then" => Arc::clone(&strings.then),
            "keys" => Arc::clone(&strings.keys),
            "from" => Arc::clone(&strings.from),
            "push" => Arc::clone(&strings.push),
            "join" => Arc::clone(&strings.join),
            "find" => Arc::clone(&strings.find),
            "sort" => Arc::clone(&strings.sort),
            "flat" => Arc::clone(&strings.flat),
            "fill" => Arc::clone(&strings.fill),
            "some" => Arc::clone(&strings.some),
            "Null" => Arc::clone(&strings.null_upper),
            "even" => Arc::clone(&strings.even),
            _ => return None,
        },
        5 => match text {
            "false" => Arc::clone(&strings.false_),
            "value" => Arc::clone(&strings.value),
            "Array" => Arc::clone(&strings.array),
            "Error" => Arc::clone(&strings.error),
            "apply" => Arc::clone(&strings.apply),
            "throw" => Arc::clone(&strings.throw),
            "cause" => Arc::clone(&strings.cause),
            "stack" => Arc::clone(&strings.stack),
            "flags" => Arc::clone(&strings.flags),
            "input" => Arc::clone(&strings.input),
            "index" => Arc::clone(&strings.index),
            "catch" => Arc::clone(&strings.catch),
            "parse" => Arc::clone(&strings.parse),
            "shift" => Arc::clone(&strings.shift),
            "slice" => Arc::clone(&strings.slice),
            "every" => Arc::clone(&strings.every),
            _ => return None,
        },
        6 => match text {
            "length" => Arc::clone(&strings.length),
            "String" => Arc::clone(&strings.string),
            "Object" => Arc::clone(&strings.object),
            "Number" => Arc::clone(&strings.number),
            "Symbol" => Arc::clone(&strings.symbol),
            "BigInt" => Arc::clone(&strings.bigint),
            "RegExp" => Arc::clone(&strings.regexp),
            "revoke" => Arc::clone(&strings.revoke),
            "return" => Arc::clone(&strings.return_),
            "source" => Arc::clone(&strings.source),
            "global" => Arc::clone(&strings.global),
            "groups" => Arc::clone(&strings.groups),
            "dotAll" => Arc::clone(&strings.dot_all),
            "sticky" => Arc::clone(&strings.sticky),
            "buffer" => Arc::clone(&strings.buffer),
            "values" => Arc::clone(&strings.values),
            "delete" => Arc::clone(&strings.delete),
            "splice" => Arc::clone(&strings.splice),
            "filter" => Arc::clone(&strings.filter),
            "reduce" => Arc::clone(&strings.reduce),
            "concat" => Arc::clone(&strings.concat),
            "Module" => Arc::clone(&strings.module),
            "__type" => Arc::clone(&strings.type_marker),
            "__keys" => Arc::clone(&strings.keys_marker),
            _ => return None,
        },
        7 => match text {
            "valueOf" => Arc::clone(&strings.value_of),
            "Boolean" => Arc::clone(&strings.boolean),
            "Promise" => Arc::clone(&strings.promise),
            "RawJSON" => Arc::clone(&strings.raw_json),
            "WeakMap" => Arc::clone(&strings.weak_map),
            "WeakSet" => Arc::clone(&strings.weak_set),
            "WeakRef" => Arc::clone(&strings.weak_ref),
            "pending" => Arc::clone(&strings.pending),
            "message" => Arc::clone(&strings.message),
            "indices" => Arc::clone(&strings.indices),
            "unicode" => Arc::clone(&strings.unicode),
            "finally" => Arc::clone(&strings.finally),
            "entries" => Arc::clone(&strings.entries),
            "isArray" => Arc::clone(&strings.is_array),
            "unshift" => Arc::clone(&strings.unshift),
            "indexOf" => Arc::clone(&strings.index_of),
            "forEach" => Arc::clone(&strings.for_each),
            "reverse" => Arc::clone(&strings.reverse),
            "flatMap" => Arc::clone(&strings.flat_map),
            "default" => Arc::clone(&strings.default),
            _ => return None,
        },
        8 => match text {
            "toString" => Arc::clone(&strings.to_string),
            "Function" => Arc::clone(&strings.function),
            "ecma:map" => Arc::clone(&strings.ecma_map),
            "ecma:set" => Arc::clone(&strings.ecma_set),
            "Iterator" => Arc::clone(&strings.iterator),
            "DataView" => Arc::clone(&strings.data_view),
            "rejected" => Arc::clone(&strings.rejected),
            "Symbol()" => Arc::clone(&strings.empty_symbol),
            "includes" => Arc::clone(&strings.includes),
            "writable" => Arc::clone(&strings.writable),
            _ => return None,
        },
        9 => match text {
            "undefined" => Arc::clone(&strings.undefined),
            "prototype" => Arc::clone(&strings.prototype),
            "anonymous" => Arc::clone(&strings.anonymous),
            "ecma:date" => Arc::clone(&strings.ecma_date),
            "ecma:json" => Arc::clone(&strings.ecma_json),
            "ecma:math" => Arc::clone(&strings.ecma_math),
            "ecma:intl" => Arc::clone(&strings.ecma_intl),
            "fulfilled" => Arc::clone(&strings.fulfilled),
            "timed-out" => Arc::clone(&strings.timed_out),
            "not-equal" => Arc::clone(&strings.not_equal),
            "multiline" => Arc::clone(&strings.multiline),
            "lastIndex" => Arc::clone(&strings.last_index),
            "stringify" => Arc::clone(&strings.stringify),
            "Undefined" => Arc::clone(&strings.undefined_upper),
            "__proto__" => Arc::clone(&strings.proto_marker),
            "__nonenum" => Arc::clone(&strings.nonenum_marker),
            _ => return None,
        },
        10 => match text {
            "ecma:array" => Arc::clone(&strings.ecma_array),
            "ecma:error" => Arc::clone(&strings.ecma_error),
            "ecma:value" => Arc::clone(&strings.ecma_value),
            "ecma:proxy" => Arc::clone(&strings.ecma_proxy),
            "hasIndices" => Arc::clone(&strings.has_indices),
            "ignoreCase" => Arc::clone(&strings.ignore_case),
            "byteLength" => Arc::clone(&strings.byte_length),
            "byteOffset" => Arc::clone(&strings.byte_offset),
            "copyWithin" => Arc::clone(&strings.copy_within),
            "enumerable" => Arc::clone(&strings.enumerable),
            "TypedArray" => Arc::clone(&strings.typed_array),
            _ => return None,
        },
        11 => match text {
            "constructor" => Arc::clone(&strings.constructor),
            "ecma:object" => Arc::clone(&strings.ecma_object),
            "ecma:number" => Arc::clone(&strings.ecma_number),
            "ecma:string" => Arc::clone(&strings.ecma_string),
            "ecma:symbol" => Arc::clone(&strings.ecma_symbol),
            "ecma:bigint" => Arc::clone(&strings.ecma_bigint),
            "ecma:regexp" => Arc::clone(&strings.ecma_regexp),
            "ecma:global" => Arc::clone(&strings.ecma_global),
            "ArrayBuffer" => Arc::clone(&strings.array_buffer),
            "unicodeSets" => Arc::clone(&strings.unicode_sets),
            "lastIndexOf" => Arc::clone(&strings.last_index_of),
            "__primitive" => Arc::clone(&strings.primitive_marker),
            _ => return None,
        },
        12 => match text {
            "[object Map]" => Arc::clone(&strings.map_tag),
            "[object Set]" => Arc::clone(&strings.set_tag),
            "ecma:boolean" => Arc::clone(&strings.ecma_boolean),
            "ecma:reflect" => Arc::clone(&strings.ecma_reflect),
            "ecma:promise" => Arc::clone(&strings.ecma_promise),
            "ecma:atomics" => Arc::clone(&strings.ecma_atomics),
            "ecma:weakmap" => Arc::clone(&strings.ecma_weakmap),
            "ecma:weakset" => Arc::clone(&strings.ecma_weakset),
            "ecma:weakref" => Arc::clone(&strings.ecma_weakref),
            "Invalid Date" => Arc::clone(&strings.invalid_date),
            "configurable" => Arc::clone(&strings.configurable),
            _ => return None,
        },
        13 => match text {
            "[object Null]" => Arc::clone(&strings.null_tag),
            "[object Date]" => Arc::clone(&strings.date_tag),
            "ecma:function" => Arc::clone(&strings.ecma_function),
            "ecma:iterator" => Arc::clone(&strings.ecma_iterator),
            "ecma:dataview" => Arc::clone(&strings.ecma_dataview),
            "ArrayIterator" => Arc::clone(&strings.array_iterator),
            "isPrototypeOf" => Arc::clone(&strings.is_prototype_of),
            _ => return None,
        },
        14 => match text {
            "[object Array]" => Arc::clone(&strings.array_tag),
            "ecma:generator" => Arc::clone(&strings.ecma_generator),
            "ecma:int8array" => Arc::clone(&strings.ecma_int8array),
            "AggregateError" => Arc::clone(&strings.aggregate_error),
            "hasOwnProperty" => Arc::clone(&strings.has_own_property),
            "toLocaleString" => Arc::clone(&strings.to_locale_string),
            _ => return None,
        },
        15 => match text {
            "[object Object]" => Arc::clone(&strings.object_tag),
            "[object Number]" => Arc::clone(&strings.number_tag),
            "[object BigInt]" => Arc::clone(&strings.bigint_tag),
            "[object String]" => Arc::clone(&strings.string_tag),
            "[object Symbol]" => Arc::clone(&strings.symbol_tag),
            "[object RegExp]" => Arc::clone(&strings.regexp_tag),
            "ecma:typedarray" => Arc::clone(&strings.ecma_typedarray),
            "ecma:globalThis" => Arc::clone(&strings.ecma_global_this),
            "ecma:fixedarray" => Arc::clone(&strings.ecma_fixedarray),
            "ecma:uint8array" => Arc::clone(&strings.ecma_uint8array),
            "ecma:int16array" => Arc::clone(&strings.ecma_int16array),
            "ecma:int32array" => Arc::clone(&strings.ecma_int32array),
            _ => return None,
        },
        16 => match text {
            "[object Boolean]" => Arc::clone(&strings.boolean_tag),
            "[object Promise]" => Arc::clone(&strings.promise_tag),
            "ecma:arraybuffer" => Arc::clone(&strings.ecma_arraybuffer),
            "ecma:intl/locale" => Arc::clone(&strings.ecma_intl_locale),
            "ecma:uint16array" => Arc::clone(&strings.ecma_uint16array),
            "ecma:uint32array" => Arc::clone(&strings.ecma_uint32array),
            _ => return None,
        },
        17 => match text {
            "[object Function]" => Arc::clone(&strings.function_tag),
            "ecma:uint8clamped" => Arc::clone(&strings.ecma_uint8clamped),
            "ecma:float32array" => Arc::clone(&strings.ecma_float32array),
            "ecma:float64array" => Arc::clone(&strings.ecma_float64array),
            "SharedArrayBuffer" => Arc::clone(&strings.shared_array_buffer),
            _ => return None,
        },
        18 => match text {
            "[object Undefined]" => Arc::clone(&strings.undefined_tag),
            "ecma:intl/collator" => Arc::clone(&strings.ecma_intl_collator),
            "ecma:intl/timezone" => Arc::clone(&strings.ecma_intl_timezone),
            "ecma:bigint64array" => Arc::clone(&strings.ecma_bigint64array),
            _ => return None,
        },
        19 => match text {
            "[object TypedArray]" => Arc::clone(&strings.typed_array_tag),
            "ecma:intl/segmenter" => Arc::clone(&strings.ecma_intl_segmenter),
            "ecma:biguint64array" => Arc::clone(&strings.ecma_biguint64array),
            _ => return None,
        },
        20 => match text {
            "[object ArrayBuffer]" => Arc::clone(&strings.array_buffer_tag),
            "ecma:intl/listformat" => Arc::clone(&strings.ecma_intl_listformat),
            "FinalizationRegistry" => Arc::clone(&strings.finalization_registry),
            "propertyIsEnumerable" => Arc::clone(&strings.property_is_enumerable),
            _ => return None,
        },
        21 => match text {
            "ecma:structured-clone" => Arc::clone(&strings.ecma_structured_clone),
            "ecma:intl/pluralrules" => Arc::clone(&strings.ecma_intl_pluralrules),
            _ => return None,
        },
        22 => match text {
            "ecma:sharedarraybuffer" => Arc::clone(&strings.ecma_sharedarraybuffer),
            "ecma:intl/numberformat" => Arc::clone(&strings.ecma_intl_numberformat),
            "ecma:intl/displaynames" => Arc::clone(&strings.ecma_intl_displaynames),
            _ => return None,
        },
        24 => match text {
            "ecma:intl/datetimeformat" => Arc::clone(&strings.ecma_intl_datetimeformat),
            "ecma:intl/durationformat" => Arc::clone(&strings.ecma_intl_durationformat),
            _ => return None,
        },
        25 => match text {
            "ecma:finalizationregistry" => Arc::clone(&strings.ecma_finalizationregistry),
            _ => return None,
        },
        26 => match text {
            "ecma:finalization-registry" => Arc::clone(&strings.ecma_finalization_registry),
            _ => return None,
        },
        28 => match text {
            "ecma:intl/relativetimeformat" => Arc::clone(&strings.ecma_intl_relativetimeformat),
            _ => return None,
        },
        29 => match text {
            "function () { [native code] }" => Arc::clone(&strings.native_function),
            _ => return None,
        },
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
fn metadata_key_len(prefix: usize, key: usize, suffix: usize) -> usize {
    prefix
        .checked_add(key)
        .and_then(|len| len.checked_add(suffix))
        .expect("property key length overflow")
}

#[inline]
pub fn with_prefixed_property_key<R>(prefix: &str, key: &str, f: impl FnOnce(&str) -> R) -> R {
    const INLINE_KEY_BYTES: usize = 128;
    let len = metadata_key_len(prefix.len(), key.len(), 0);
    if len <= INLINE_KEY_BYTES {
        let mut buf = std::mem::MaybeUninit::<[u8; INLINE_KEY_BYTES]>::uninit();
        // The capacity check covers both copies. Only the initialized prefix is
        // exposed, and concatenating valid str inputs preserves UTF-8 validity.
        let text = unsafe {
            let ptr = buf.as_mut_ptr().cast::<u8>();
            std::ptr::copy_nonoverlapping(prefix.as_ptr(), ptr, prefix.len());
            std::ptr::copy_nonoverlapping(key.as_ptr(), ptr.add(prefix.len()), key.len());
            std::str::from_utf8_unchecked(std::slice::from_raw_parts(ptr, len))
        };
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
    let len = metadata_key_len(prefix.len(), key.len(), suffix.len());
    if len <= INLINE_KEY_BYTES {
        let mut buf = std::mem::MaybeUninit::<[u8; INLINE_KEY_BYTES]>::uninit();
        // All three ranges fit, are disjoint, and initialize every byte read.
        // Their inputs are valid UTF-8, including at concatenation boundaries.
        let text = unsafe {
            let ptr = buf.as_mut_ptr().cast::<u8>();
            std::ptr::copy_nonoverlapping(prefix.as_ptr(), ptr, prefix.len());
            std::ptr::copy_nonoverlapping(key.as_ptr(), ptr.add(prefix.len()), key.len());
            std::ptr::copy_nonoverlapping(
                suffix.as_ptr(), ptr.add(prefix.len() + key.len()), suffix.len(),
            );
            std::str::from_utf8_unchecked(std::slice::from_raw_parts(ptr, len))
        };
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

pub use vybe_runtime::string_concat::{concat2_arc, concat3_arc};

#[cfg(test)]
mod direct_arc_concat_tests {
    use super::*;
    use std::hint::black_box;
    use std::time::{Duration, Instant};

    fn previous_concat2(left: &str, right: &str) -> Arc<str> {
        let mut out = String::with_capacity(left.len() + right.len());
        out.push_str(left);
        out.push_str(right);
        Arc::<str>::from(out)
    }

    fn previous_concat3(left: &str, middle: &str, right: &str) -> Arc<str> {
        let mut out = String::with_capacity(left.len() + middle.len() + right.len());
        out.push_str(left);
        out.push_str(middle);
        out.push_str(right);
        Arc::<str>::from(out)
    }

    #[test]
    fn direct_concatenation_preserves_utf8_empty_nul_and_boundaries() {
        for (left, middle, right) in [
            ("", "", ""),
            ("", "value", ""),
            ("prefix", "", "suffix"),
            ("[object ", "Custom", "]"),
            ("\u{e9}", "\u{1f680}", "\u{4e2d}"),
            ("a\0", "\0b", "\0"),
        ] {
            assert_eq!(concat2_arc(left, right), previous_concat2(left, right));
            assert_eq!(concat3_arc(left, middle, right), previous_concat3(left, middle, right));
        }
        for len in [1, 7, 15, 31, 63, 127, 128, 129, 255, 1024, 4096] {
            let input = "x".repeat(len);
            assert_eq!(concat2_arc(&input, &input), previous_concat2(&input, &input));
            assert_eq!(concat3_arc(&input, &input, &input), previous_concat3(&input, &input, &input));
        }
    }

    #[test]
    fn concatenation_keeps_sources_and_shared_results_independent() {
        let source: Arc<str> = Arc::from("shared\u{1f680}");
        let result = concat3_arc(&source, &source, &source);
        assert_eq!(Arc::strong_count(&source), 1);
        assert_eq!(result.as_ref(), "shared\u{1f680}shared\u{1f680}shared\u{1f680}");
        let alias = result.clone();
        assert!(Arc::ptr_eq(&result, &alias));
        drop(result);
        assert_eq!(alias.as_ref(), "shared\u{1f680}shared\u{1f680}shared\u{1f680}");
        assert_eq!(source.as_ref(), "shared\u{1f680}");
    }

    #[test]
    fn every_empty_part_combination_matches_previous_concatenation() {
        for mask in 0..8 {
            let left = if mask & 1 == 0 { "" } else { "\u{e9}" };
            let middle = if mask & 2 == 0 { "" } else { "\u{1f680}" };
            let right = if mask & 4 == 0 { "" } else { "x\0y" };
            assert_eq!(concat2_arc(left, right), previous_concat2(left, right));
            assert_eq!(concat3_arc(left, middle, right), previous_concat3(left, middle, right));
        }
    }

    #[test]
    #[ignore = "paired complete native concatenation benchmark, not application timing"]
    fn paired_direct_arc_concat_benchmark() {
        for len in [0, 8, 64, 512, 4096] {
            let input = "x".repeat(len);
            for parts in [2, 3] {
                let timed = |direct: bool| {
                    let start = Instant::now();
                    for _ in 0..10_000 {
                        let result = match (direct, parts) {
                            (false, 2) => previous_concat2(black_box(&input), black_box("suffix")),
                            (true, 2) => concat2_arc(black_box(&input), black_box("suffix")),
                            (false, _) => previous_concat3(black_box("[object "), black_box(&input), black_box("]")),
                            (true, _) => concat3_arc(black_box("[object "), black_box(&input), black_box("]")),
                        };
                        black_box(result);
                    }
                    start.elapsed()
                };
                let mut previous = Duration::ZERO;
                let mut direct = Duration::ZERO;
                for round in 0..8 {
                    if round % 2 == 0 {
                        previous += timed(false);
                        direct += timed(true);
                    } else {
                        direct += timed(true);
                        previous += timed(false);
                    }
                }
                eprintln!("arc concat input={len} parts={parts}: previous={previous:?}, direct={direct:?}");
            }
        }
    }
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
    let bytes = key.as_bytes();
    if bytes.is_empty() || bytes.len() > 10 {
        return None;
    }
    let first = bytes[0];
    if first == b'0' {
        return (bytes.len() == 1).then_some(0);
    }
    if !first.is_ascii_digit() {
        return None;
    }
    // At most ten decimal digits fit below 10^10, so accumulation cannot
    // overflow u64. Canonical indexes exclude 2^32 - 1 itself.
    let mut value = (first - b'0') as u64;
    for &byte in &bytes[1..] {
        if !byte.is_ascii_digit() {
            return None;
        }
        value = value * 10 + (byte - b'0') as u64;
    }
    (value < u32::MAX as u64).then_some(value as u32)
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

#[cfg(test)]
mod bounded_array_index_parser_tests {
    use super::*;
    use std::hint::black_box;
    use std::time::{Duration, Instant};

    fn previous_array_index_key(key: &str) -> Option<u32> {
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
    

    #[test]
    fn canonical_boundaries_and_invalid_forms_are_preserved() {
        for (text, expected) in [
            ("", None), ("0", Some(0)), ("1", Some(1)), ("255", Some(255)),
            ("4294967294", Some(u32::MAX - 1)), ("4294967295", None),
            ("4294967296", None), ("9999999999", None), ("10000000000", None),
            ("00", None), ("01", None), ("00000000001", None),
            ("-0", None), ("-1", None), ("+1", None), (" 1", None), ("1 ", None),
            ("1.0", None), ("1e0", None), ("1_0", None), ("1\0", None),
            ("field", None), ("\u{ff11}", None), ("\u{1f680}", None),
        ] {
            assert_eq!(canonical_array_index_key(text), expected, "{text:?}");
            assert_eq!(canonical_array_index_key(text), previous_array_index_key(text), "{text:?}");
        }
        let huge = "9".repeat(4096);
        assert_eq!(canonical_array_index_key(&huge), None);
        assert_eq!(canonical_array_index_key(&huge), previous_array_index_key(&huge));
    }

    #[test]
    fn broad_decimal_and_malformed_corpus_matches_previous() {
        for number in 0..100_000u32 {
            let text = number.to_string();
            assert_eq!(canonical_array_index_key(&text), Some(number));
            assert_eq!(canonical_array_index_key(&text), previous_array_index_key(&text));
        }
        let mut state = 0x9e37_79b9u32;
        let alphabet = b"0123456789+-e. x";
        for _ in 0..20_000 {
            state = state.wrapping_mul(1664525).wrapping_add(1013904223);
            let length = (state % 32) as usize;
            let mut text = String::with_capacity(length);
            for _ in 0..length {
                state = state.wrapping_mul(1664525).wrapping_add(1013904223);
                text.push(alphabet[(state as usize) % alphabet.len()] as char);
            }
            assert_eq!(canonical_array_index_key(&text), previous_array_index_key(&text), "{text:?}");
        }
    }

    #[test]
    #[ignore = "paired shared index classification benchmark, not lookup or guest timing"]
    fn paired_bounded_parser_benchmark() {
        for (name, keys) in [
            ("small", vec!["0", "1", "7", "17", "255", "1024"]),
            ("boundary", vec!["4294967294", "4294967295", "9999999999", "10000000000"]),
            ("ordinary", vec!["a", "field", "length", "__descriptor", "applicationField"]),
            ("invalid", vec!["01", "-0", "1e0", "1\0", "999999999999999999999999"]),
        ] {
            let timed = |current: bool| {
                let start = Instant::now();
                for _ in 0..20_000 {
                    for key in &keys {
                        black_box(if current { canonical_array_index_key(black_box(key)) } else { previous_array_index_key(black_box(key)) });
                    }
                }
                start.elapsed()
            };
            let mut previous = Duration::ZERO;
            let mut current = Duration::ZERO;
            for round in 0..8 {
                if round % 2 == 0 { previous += timed(false); current += timed(true); }
                else { current += timed(true); previous += timed(false); }
            }
            eprintln!("bounded index corpus={name}: previous={previous:?}, current={current:?}");
        }
    }
}

#[cfg(test)]
mod metadata_key_buffer_tests {
    use super::*;

    #[test]
    fn metadata_length_boundaries() {
        assert_eq!(metadata_key_len(0, 0, 0), 0);
        assert_eq!(metadata_key_len(6, 122, 0), 128);
        assert_eq!(metadata_key_len(6, 120, 2), 128);
        assert_eq!(metadata_key_len(6, 121, 2), 129);
        assert_eq!(metadata_key_len(usize::MAX, 0, 0), usize::MAX);
    }

    #[test]
    #[should_panic(expected = "property key length overflow")]
    fn prefix_key_length_overflow_is_rejected() {
        metadata_key_len(usize::MAX, 1, 0);
    }

    #[test]
    #[should_panic(expected = "property key length overflow")]
    fn suffix_length_overflow_is_rejected() {
        metadata_key_len(usize::MAX - 1, 1, 1);
    }

    #[test]
    fn initialized_key_ranges_match_concatenation() {
        let cases = [
            ("", "", ""),
            ("__get_", "field0", ""),
            ("__set_", "length", ""),
            ("\u{e9}", "\u{1f680}", "\u{4e2d}"),
            ("prefix", "embedded\0key", "suffix"),
        ];
        for (prefix, key, suffix) in cases {
            with_prefixed_property_key(prefix, key, |text| {
                assert_eq!(text, format!("{prefix}{key}"));
            });
            with_wrapped_property_key(prefix, key, suffix, |text| {
                assert_eq!(text, format!("{prefix}{key}{suffix}"));
            });
        }
        for total in [127, 128, 129, 256] {
            let key = "x".repeat(total - 8);
            with_prefixed_property_key("__get_", &key, |text| {
                assert_eq!(text, format!("__get_{key}"));
            });
            with_wrapped_property_key("__get_", &key, "[]", |text| {
                assert_eq!(text.len(), total);
                assert_eq!(text, format!("__get_{key}[]"));
            });
        }
        for len in [127, 128, 129] {
            let key = "x".repeat(len - 6);
            with_prefixed_property_key("__get_", &key, |text| {
                assert_eq!(text.len(), len);
                assert_eq!(text, format!("__get_{key}"));
            });
        }
    }

    #[test]
    #[ignore = "matched metadata-key microbenchmark"]
    fn benchmark_prefixed_key_buffer() {
        use std::hint::black_box;
        use std::time::Instant;
        let samples = ["field0", "length", "someLongerApplicationPropertyName"];
        for round in 0..4 {
            let old = || {
                let start = Instant::now();
                for _ in 0..100_000 {
                    for key in samples {
                        let prefix = black_box("__get_");
                        let key = black_box(key);
                        let len = prefix.len() + key.len();
                        let mut buf = [0u8; 128];
                        buf[..prefix.len()].copy_from_slice(prefix.as_bytes());
                        buf[prefix.len()..len].copy_from_slice(key.as_bytes());
                        black_box(std::str::from_utf8(&buf[..len]).unwrap_or(""));
                    }
                }
                start.elapsed()
            };
            let new = || {
                let start = Instant::now();
                for _ in 0..100_000 {
                    for key in samples {
                        with_prefixed_property_key(black_box("__get_"), black_box(key), |text| {
                            black_box(text);
                        });
                    }
                }
                start.elapsed()
            };
            let (before, after) = if round % 2 == 0 {
                (old(), new())
            } else {
                let after = new();
                (old(), after)
            };
            eprintln!("metadata keys round={round} original={before:?} optimized={after:?}");
        }
    }
}
