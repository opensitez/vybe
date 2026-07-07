//! Typed array constructors — element kinds, buffer views, BYTES_PER_ELEMENT.

crate::js_cases! {
    int8array_from_array_sets_values => {
        r#"const a=new Int8Array([-1,0,1]); console.log(a[0]);console.log(a[2]);"#,
        ["-1", "1"]
    };

    uint8array_from_array_sets_values => {
        r#"const a=new Uint8Array([0,128,255]); console.log(a[1]);console.log(a[2]);"#,
        ["128", "255"]
    };

    uint8clampedarray_clamps_above_255 => {
        r#"const a=new Uint8ClampedArray([300]); console.log(a[0]);"#,
        ["255"]
    };

    uint8clampedarray_clamps_below_zero => {
        r#"const a=new Uint8ClampedArray([-10]); console.log(a[0]);"#,
        ["0"]
    };

    int16array_bytes_per_element => {
        r#"console.log(Int16Array.BYTES_PER_ELEMENT);"#,
        ["2"]
    };

    uint16array_from_buffer_slice => {
        r#"const buf=new ArrayBuffer(4); const a=new Uint16Array(buf); a[0]=0xabcd; console.log(a[0].toString(16));"#,
        ["abcd"]
    };

    int32array_negative_values => {
        r#"const a=new Int32Array([-100,100]); console.log(a[0]);console.log(a[1]);"#,
        ["-100", "100"]
    };

    uint32array_max_unsigned => {
        r#"const a=new Uint32Array([4294967295]); console.log(a[0]);"#,
        ["4294967295"]
    };

    float32array_fractional_values => {
        r#"const a=new Float32Array([1.5]); console.log(a[0]);"#,
        ["1.5"]
    };

    float64array_high_precision => {
        r#"const a=new Float64Array([1.123456789]); console.log(a[0]>1.123456);"#,
        ["true"]
    };

    // Node-verified: BigInt64Array elements are BigInts, printed with `n`.
    bigint64array_bigint_elements => {
        r#"const a=new BigInt64Array([1n,-1n]); console.log(a[0]);console.log(a[1]);"#,
        ["1n", "-1n"]
    };

    // Node-verified: prints with `n`. Exact — BigInt is arbitrary
    // precision (u64::MAX round-trips via the ToBigUint64 reading).
    biguint64array_large_bigint => {
        r#"const a=new BigUint64Array([18446744073709551615n]); console.log(a[0]);"#,
        ["18446744073709551615n"]
    };

    typed_array_length_from_constructor => {
        r#"console.log(new Uint8Array(5).length);"#,
        ["5"]
    };

    typed_array_buffer_property => {
        r#"const buf=new ArrayBuffer(8); const a=new Uint8Array(buf); console.log(a.buffer===buf);"#,
        ["true"]
    };

    typed_array_byte_offset_from_slice => {
        r#"const a=new Uint8Array(new ArrayBuffer(4),1,2); console.log(a.byteOffset);console.log(a.byteLength);"#,
        ["1", "2"]
    };

    int8array_set_from_another_typed_array => {
        r#"const src=new Int8Array([1,2,3]); const dst=new Int8Array(3); dst.set(src); console.log(dst[2]);"#,
        ["3"]
    };

    uint8array_subarray_shares_buffer => {
        r#"const a=new Uint8Array([1,2,3,4]); const s=a.subarray(1,3); s[0]=9; console.log(a[1]);"#,
        ["9"]
    };

    int16array_slice_copies_elements => {
        r#"const a=new Int16Array([1,2,3]); const s=a.slice(1); s[0]=99; console.log(a[1]);"#,
        ["2"]
    };

    float32array_map_to_plain_array => {
        r#"const a=new Float32Array([2,3]); const m=Array.from(a,x=>x*2); console.log(m.join(","));"#,
        ["4,6"]
    };

    uint8array_join_with_separator => {
        r#"console.log(new Uint8Array([1,2,3]).join("-"));"#,
        ["1-2-3"]
    };

    int32array_index_of_finds_value => {
        r#"console.log(new Int32Array([5,6,7]).indexOf(6));"#,
        ["1"]
    };

    uint8array_includes_checks_membership => {
        r#"console.log(new Uint8Array([1,2]).includes(2));"#,
        ["true"]
    };

    int8array_reverse_mutates_in_place => {
        r#"const a=new Int8Array([1,2,3]); a.reverse(); console.log(a[0]);"#,
        ["3"]
    };

    uint8array_sort_default_numeric => {
        r#"const a=new Uint8Array([3,1,2]); a.sort(); console.log(a.join(","));"#,
        ["1,2,3"]
    };

    float64array_fill_sets_all => {
        r#"const a=new Float64Array(3); a.fill(2.5); console.log(a[1]);"#,
        ["2.5"]
    };

    int16array_copy_within_moves_range => {
        r#"const a=new Int16Array([1,2,3,4]); a.copyWithin(0,2); console.log(a[0]);console.log(a[1]);"#,
        ["3", "4"]
    };

    uint8array_from_hex_like_values => {
        r#"const a=Uint8Array.from([0,15,255]); console.log(a[1]);"#,
        ["15"]
    };

    int32array_of_constructor => {
        r#"console.log(Int32Array.of(1,2,3).length);"#,
        ["3"]
    };

    uint8array_from_string_maps_char_codes => {
        r#"const a=Uint8Array.from("AB",c=>c.charCodeAt(0)); console.log(a[0]);console.log(a[1]);"#,
        ["65", "66"]
    };

    bigint64array_subarray_preserves_element_type => {
        r#"const a=new BigInt64Array([1n,2n,3n]); console.log(typeof a.subarray(1)[0]);"#,
        ["bigint"]
    };

    typed_array_iterator_next_returns_value => {
        r#"const it=new Uint8Array([4,5]).values(); console.log(it.next().value);"#,
        ["4"]
    };

    typed_array_entries_yields_index_value_pairs => {
        r#"const e=new Uint8Array([7]).entries().next().value; console.log(e[0]);console.log(e[1]);"#,
        ["0", "7"]
    };

    typed_array_keys_yields_indices => {
        r#"console.log(new Uint8Array(2).keys().next().value);"#,
        ["0"]
    };

    int8array_set_with_offset => {
        r#"const dst=new Int8Array(4); dst.set([9,8],2); console.log(dst[2]);console.log(dst[3]);"#,
        ["9", "8"]
    };

    uint8array_out_of_range_get_returns_undefined => {
        r#"console.log(new Uint8Array([1])[5]);"#,
        ["undefined"]
    };

    float32array_nan_in_array => {
        r#"console.log(Number.isNaN(new Float32Array([NaN])[0]));"#,
        ["true"]
    };

    int16array_length_zero => {
        r#"console.log(new Int16Array(0).length);"#,
        ["0"]
    };

    uint8array_prototype_is_typed_array => {
        // §23.2.3.38: the `Symbol.toStringTag` getter returns the typed-array
        // name for an INSTANCE, but `undefined` for the prototype itself
        // (not a typed array). Test the instance (node-verified).
        r#"console.log(new Uint8Array(1)[Symbol.toStringTag]);"#,
        ["Uint8Array"]
    };

    shared_typed_array_from_shared_array_buffer => {
        r#"const a=new Int32Array(new SharedArrayBuffer(4)); a[0]=11; console.log(a[0]);"#,
        ["11"]
    };

    int8array_to_locale_string_joins => {
        r#"console.log(new Int8Array([1,2]).toLocaleString().includes("1"));"#,
        ["true"]
    };

    uint8clampedarray_rounds_floats => {
        r#"const a=new Uint8ClampedArray([1.4,1.6]); console.log(a[0]);console.log(a[1]);"#,
        ["1", "2"]
    };

    int32array_find_finds_element => {
        r#"console.log(new Int32Array([1,9,3]).find(x=>x>5));"#,
        ["9"]
    };

    uint8array_some_checks_predicate => {
        r#"console.log(new Uint8Array([1,2,3]).some(x=>x===2));"#,
        ["true"]
    };

    int16array_every_checks_all => {
        r#"console.log(new Int16Array([2,4,6]).every(x=>x%2===0));"#,
        ["true"]
    };

    float32array_reduce_sums => {
        r#"console.log(new Float32Array([1,2,3]).reduce((a,b)=>a+b,0));"#,
        ["6"]
    };

    uint8array_filter_creates_plain_array => {
        r#"const r=new Uint8Array([1,2,3]).filter(x=>x>1); console.log(Array.isArray(r));console.log(r.length);"#,
        ["true", "2"]
    };
}
