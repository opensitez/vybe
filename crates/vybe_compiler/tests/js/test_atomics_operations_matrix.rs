//! Atomics operations on SharedArrayBuffer — RMW, compareExchange, load/store.

crate::js_cases! {
    atomics_add_returns_old_value => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=5; console.log(Atomics.add(ia,0,3));console.log(ia[0]);"#,
        ["5", "8"]
    };

    atomics_sub_returns_old_value => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=10; console.log(Atomics.sub(ia,0,4));console.log(ia[0]);"#,
        ["10", "6"]
    };

    atomics_and_bitwise_and => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=0b1111; Atomics.and(ia,0,0b1010); console.log(ia[0]);"#,
        ["10"]
    };

    atomics_or_bitwise_or => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=0b1000; Atomics.or(ia,0,0b0011); console.log(ia[0]);"#,
        ["11"]
    };

    atomics_xor_bitwise_xor => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=0b1100; Atomics.xor(ia,0,0b1010); console.log(ia[0]);"#,
        ["6"]
    };

    atomics_exchange_replaces_with_new_value => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=7; console.log(Atomics.exchange(ia,0,99));console.log(ia[0]);"#,
        ["7", "99"]
    };

    atomics_compare_exchange_succeeds_when_equal => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=5; console.log(Atomics.compareExchange(ia,0,5,9));console.log(ia[0]);"#,
        ["5", "9"]
    };

    atomics_compare_exchange_fails_when_not_equal => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=5; console.log(Atomics.compareExchange(ia,0,4,9));console.log(ia[0]);"#,
        ["5", "5"]
    };

    atomics_load_reads_current_value => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=42; console.log(Atomics.load(ia,0));"#,
        ["42"]
    };

    atomics_store_writes_and_returns_value => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); console.log(Atomics.store(ia,0,13));console.log(ia[0]);"#,
        ["13", "13"]
    };

    atomics_is_lock_free_size_one => {
        r#"console.log(Atomics.isLockFree(1));"#,
        ["true"]
    };

    atomics_is_lock_free_size_two => {
        r#"console.log(typeof Atomics.isLockFree(2));"#,
        ["boolean"]
    };

    atomics_add_on_bigint64_array => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); ia[0]=10n; console.log(Atomics.add(ia,0,5n));console.log(ia[0]);"#,
        ["10", "15"]
    };

    atomics_sub_on_bigint64_array => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); ia[0]=20n; Atomics.sub(ia,0,3n); console.log(ia[0]);"#,
        ["17"]
    };

    atomics_and_on_bigint64_array => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); ia[0]=7n; Atomics.and(ia,0,3n); console.log(ia[0]);"#,
        ["3"]
    };

    atomics_or_on_bigint64_array => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); ia[0]=1n; Atomics.or(ia,0,6n); console.log(ia[0]);"#,
        ["7"]
    };

    atomics_xor_on_bigint64_array => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); ia[0]=5n; Atomics.xor(ia,0,3n); console.log(ia[0]);"#,
        ["6"]
    };

    atomics_exchange_on_bigint64_array => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); ia[0]=4n; console.log(Atomics.exchange(ia,0,8n));"#,
        ["4"]
    };

    atomics_compare_exchange_on_bigint64_array => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); ia[0]=2n; console.log(Atomics.compareExchange(ia,0,2n,5n));console.log(ia[0]);"#,
        ["2", "5"]
    };

    atomics_load_on_bigint64_array => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); ia[0]=99n; console.log(Atomics.load(ia,0));"#,
        ["99"]
    };

    atomics_store_on_bigint64_array => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); console.log(Atomics.store(ia,0,11n));"#,
        ["11"]
    };

    atomics_operations_on_uint32_array => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Uint32Array(sab); Atomics.store(ia,0,100); console.log(Atomics.add(ia,0,50));console.log(ia[0]);"#,
        ["100", "150"]
    };

    atomics_out_of_bounds_throws => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); try{Atomics.load(ia,1);}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    atomics_on_non_shared_buffer_throws => {
        r#"const ia=new Int32Array(1); try{Atomics.load(ia,0);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    atomics_add_negative_delta => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=3; Atomics.add(ia,0,-5); console.log(ia[0]);"#,
        ["-2"]
    };

    atomics_compare_exchange_with_expected_after_race_pattern => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=1; Atomics.store(ia,0,2); console.log(Atomics.compareExchange(ia,0,1,9));console.log(ia[0]);"#,
        ["2", "2"]
    };

    atomics_notify_no_waiters_returns_zero => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); console.log(Atomics.notify(ia,0,1));"#,
        ["0"]
    };

    atomics_wait_on_non_equal_returns_not_equal => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=1; console.log(Atomics.wait(ia,0,2,0));"#,
        ["not-equal"]
    };

    atomics_wait_timeout_returns_timed_out => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=5; console.log(Atomics.wait(ia,0,5,0));"#,
        ["timed-out"]
    };

    atomics_store_then_load_consistent => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); Atomics.store(ia,0,77); console.log(Atomics.load(ia,0));"#,
        ["77"]
    };

    atomics_xor_clears_bits => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=0b1111; Atomics.xor(ia,0,0b1111); console.log(ia[0]);"#,
        ["0"]
    };

    atomics_and_clears_high_bits => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=255; Atomics.and(ia,0,15); console.log(ia[0]);"#,
        ["15"]
    };

    atomics_or_sets_bits => {
        r#"const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=1; Atomics.or(ia,0,8); console.log(ia[0]);"#,
        ["9"]
    };

    atomics_multiple_indices_independent => {
        r#"const sab=new SharedArrayBuffer(8); const ia=new Int32Array(sab); Atomics.store(ia,0,1); Atomics.store(ia,1,2); console.log(Atomics.load(ia,0));console.log(Atomics.load(ia,1));"#,
        ["1", "2"]
    };
}
