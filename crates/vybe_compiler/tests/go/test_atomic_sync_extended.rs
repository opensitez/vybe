//! sync/atomic package: Load, Store, Add, Swap, CompareAndSwap, and Value.
//!
//! Distinct from `test_sync_package.rs` (Mutex, RWMutex, WaitGroup, Once, Pool, Map).
//! Covers function-style atomics (`atomic.LoadInt64`, `atomic.AddInt64`, …) and typed
//! `atomic.Int64` / `atomic.Value` method forms where the frontend accepts them.


go_run_cases! {
    // ── Load / Store: int64 ─────────────────────────────────────────────────
    load_int64_zero_value_defaults_to_zero => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; fmt.Println(atomic.LoadInt64(&n)) }",
        vec!["0"]
    ),
    store_int64_then_load_roundtrip => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 42); fmt.Println(atomic.LoadInt64(&n)) }",
        vec!["42"]
    ),
    store_int64_overwrites_previous => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 1); atomic.StoreInt64(&n, 99); fmt.Println(atomic.LoadInt64(&n)) }",
        vec!["99"]
    ),

    // ── Load / Store: int32, uint32, uint64 ─────────────────────────────────
    store_int32_then_load_roundtrip => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int32; atomic.StoreInt32(&n, -7); fmt.Println(atomic.LoadInt32(&n)) }",
        vec!["-7"]
    ),
    store_uint32_then_load_roundtrip => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n uint32; atomic.StoreUint32(&n, 65535); fmt.Println(atomic.LoadUint32(&n)) }",
        vec!["65535"]
    ),
    store_uint64_then_load_roundtrip => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n uint64; atomic.StoreUint64(&n, 10000000000); fmt.Println(atomic.LoadUint64(&n)) }",
        vec!["10000000000"]
    ),

    // ── Add ─────────────────────────────────────────────────────────────────
    add_int64_returns_new_value_and_updates => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 10); fmt.Println(atomic.AddInt64(&n, 5)); fmt.Println(atomic.LoadInt64(&n)) }",
        vec!["15", "15"]
    ),
    add_int64_negative_delta_decrements => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 20); fmt.Println(atomic.AddInt64(&n, -8)); fmt.Println(atomic.LoadInt64(&n)) }",
        vec!["12", "12"]
    ),
    add_int32_increments_from_zero => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int32; fmt.Println(atomic.AddInt32(&n, 3)); fmt.Println(atomic.LoadInt32(&n)) }",
        vec!["3", "3"]
    ),
    add_uint32_increments_unsigned => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n uint32; atomic.StoreUint32(&n, 100); fmt.Println(atomic.AddUint32(&n, 50)); fmt.Println(atomic.LoadUint32(&n)) }",
        vec!["150", "150"]
    ),
    add_uint64_large_counter => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n uint64; atomic.StoreUint64(&n, 9000000000); fmt.Println(atomic.AddUint64(&n, 1)); fmt.Println(atomic.LoadUint64(&n)) }",
        vec!["9000000001", "9000000001"]
    ),
    add_int64_sequential_increments => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; atomic.AddInt64(&n, 1); atomic.AddInt64(&n, 2); atomic.AddInt64(&n, 3); fmt.Println(atomic.LoadInt64(&n)) }",
        vec!["6"]
    ),

    // ── Swap ────────────────────────────────────────────────────────────────
    swap_int64_returns_previous_value => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 3); fmt.Println(atomic.SwapInt64(&n, 9)) }",
        vec!["3"]
    ),
    swap_int64_leaves_new_value_in_memory => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 3); atomic.SwapInt64(&n, 9); fmt.Println(atomic.LoadInt64(&n)) }",
        vec!["9"]
    ),
    swap_int32_replaces_atomically => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int32; atomic.StoreInt32(&n, 11); fmt.Println(atomic.SwapInt32(&n, 22)); fmt.Println(atomic.LoadInt32(&n)) }",
        vec!["11", "22"]
    ),
    swap_uint64_replaces_unsigned => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n uint64; atomic.StoreUint64(&n, 5); fmt.Println(atomic.SwapUint64(&n, 8)); fmt.Println(atomic.LoadUint64(&n)) }",
        vec!["5", "8"]
    ),

    // ── CompareAndSwap ──────────────────────────────────────────────────────
    compare_and_swap_int64_succeeds_when_expected_matches => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 1); fmt.Println(atomic.CompareAndSwapInt64(&n, 1, 2)); fmt.Println(atomic.LoadInt64(&n)) }",
        vec!["true", "2"]
    ),
    compare_and_swap_int64_fails_when_expected_mismatches => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 1); fmt.Println(atomic.CompareAndSwapInt64(&n, 9, 2)); fmt.Println(atomic.LoadInt64(&n)) }",
        vec!["false", "1"]
    ),
    compare_and_swap_int32_success_updates => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int32; atomic.StoreInt32(&n, 7); fmt.Println(atomic.CompareAndSwapInt32(&n, 7, 14)); fmt.Println(atomic.LoadInt32(&n)) }",
        vec!["true", "14"]
    ),
    compare_and_swap_int32_failure_preserves => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n int32; atomic.StoreInt32(&n, 7); fmt.Println(atomic.CompareAndSwapInt32(&n, 8, 14)); fmt.Println(atomic.LoadInt32(&n)) }",
        vec!["false", "7"]
    ),
    compare_and_swap_uint32_unsigned_values => (
        "package main; import \"fmt\"; import \"sync/atomic\"; func main() { var n uint32; atomic.StoreUint32(&n, 100); fmt.Println(atomic.CompareAndSwapUint32(&n, 100, 200)); fmt.Println(atomic.LoadUint32(&n)) }",
        vec!["true", "200"]
    ),
}

go_compile_cases! {
    // ── concurrent / goroutine patterns ───────────────────────────────────
    add_int64_goroutines_increment_shared_counter => "package main; import \"sync/atomic\"; func main() { var n int64; for i := 0; i < 5; i++ { go func() { atomic.AddInt64(&n, 1) }() }; _ = atomic.LoadInt64(&n) }",
    add_int32_concurrent_increments_compile => "package main; import \"sync/atomic\"; func main() { var n int32; go func() { atomic.AddInt32(&n, 1) }(); go func() { atomic.AddInt32(&n, 2) }(); _ = atomic.LoadInt32(&n) }",
    swap_int64_concurrent_with_load => "package main; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 1); go func() { atomic.SwapInt64(&n, 2) }(); _ = atomic.LoadInt64(&n) }",
    compare_and_swap_int64_spin_retry_pattern => "package main; import \"sync/atomic\"; func main() { var n int64; atomic.StoreInt64(&n, 0); for !atomic.CompareAndSwapInt64(&n, 0, 1) { }; _ = n }",
    store_load_int64_defer_after_goroutine => "package main; import \"sync/atomic\"; func main() { var n int64; go func() { atomic.StoreInt64(&n, 5) }(); _ = atomic.LoadInt64(&n) }",
    add_uint64_concurrent_counter => "package main; import \"sync/atomic\"; func main() { var n uint64; for i := 0; i < 3; i++ { go func() { atomic.AddUint64(&n, 1) }() }; _ = atomic.LoadUint64(&n) }",

    // ── typed atomic.Int64 methods (Go 1.19+) ─────────────────────────────
    typed_int64_store_load_methods => "package main; import \"sync/atomic\"; func main() { var v atomic.Int64; v.Store(5); _ = v.Load() }",
    typed_int64_add_method => "package main; import \"sync/atomic\"; func main() { var v atomic.Int64; v.Add(3); _ = v.Load() }",
    typed_int64_swap_method => "package main; import \"sync/atomic\"; func main() { var v atomic.Int64; v.Store(1); _ = v.Swap(9) }",
    typed_int64_compare_and_swap_method => "package main; import \"sync/atomic\"; func main() { var v atomic.Int64; v.Store(1); _ = v.CompareAndSwap(1, 2) }",
    typed_uint64_add_and_load => "package main; import \"sync/atomic\"; func main() { var v atomic.Uint64; v.Add(10); _ = v.Load() }",

    // ── atomic.Value Store / Load ───────────────────────────────────────────
    atomic_value_store_load_string => "package main; import \"sync/atomic\"; func main() { var v atomic.Value; v.Store(\"hello\"); _ = v.Load().(string) }",
    atomic_value_store_load_int => "package main; import \"sync/atomic\"; func main() { var v atomic.Value; v.Store(42); _ = v.Load().(int) }",
    atomic_value_overwrite_replaces_stored => "package main; import \"sync/atomic\"; func main() { var v atomic.Value; v.Store(1); v.Store(2); _ = v.Load().(int) }",
    atomic_value_store_struct_pointer => "package main; import \"sync/atomic\"; type cfg struct { Port int }; func main() { var v atomic.Value; v.Store(&cfg{Port: 8080}); _ = v.Load().(*cfg) }",
}
