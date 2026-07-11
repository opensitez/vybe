//! sync.Map, sync.Once, sync.Pool, and sync.Cond extended patterns.
//! Distinct from `test_sync_package.rs` (basic Mutex/Map/Once/Pool) and
//! `test_atomic_sync_extended.rs` (sync/atomic primitives).

go_run_cases! {
    sync_map_int_key_store_load =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(1, \"one\"); v, ok := m.Load(1); fmt.Println(v.(string)); fmt.Println(ok) }", vec!["one", "true"]),
    sync_map_int_key_load_missing =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; _, ok := m.Load(99); fmt.Println(ok) }", vec!["false"]),
    sync_map_struct_value_roundtrip =>
        ("package main; import \"fmt\"; import \"sync\"; type item struct { n int }; func main() { var m sync.Map; m.Store(\"k\", item{n: 8}); v, ok := m.Load(\"k\"); fmt.Println(v.(item).n); fmt.Println(ok) }", vec!["8", "true"]),
    sync_map_load_or_store_keeps_first =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"a\", 10); actual, loaded := m.LoadOrStore(\"a\", 99); fmt.Println(actual.(int)); fmt.Println(loaded) }", vec!["10", "true"]),
    sync_map_load_or_store_inserts_new =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; actual, loaded := m.LoadOrStore(\"b\", 3); fmt.Println(actual.(int)); fmt.Println(loaded) }", vec!["3", "false"]),
    sync_map_load_and_delete_returns_value =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"x\", 5); v, ok := m.LoadAndDelete(\"x\"); _, still := m.Load(\"x\"); fmt.Println(v.(int)); fmt.Println(ok); fmt.Println(still) }", vec!["5", "true", "false"]),
    sync_map_load_and_delete_missing =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; _, ok := m.LoadAndDelete(\"z\"); fmt.Println(ok) }", vec!["false"]),
    sync_map_delete_then_load =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"d\", 1); m.Delete(\"d\"); _, ok := m.Load(\"d\"); fmt.Println(ok) }", vec!["false"]),
    sync_map_range_sums_int_values =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(1, 10); m.Store(2, 20); sum := 0; m.Range(func(k, v interface{}) bool { sum += v.(int); return true }); fmt.Println(sum) }", vec!["30"]),
    sync_map_range_stops_early =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(1, 1); m.Store(2, 2); count := 0; m.Range(func(k, v interface{}) bool { count++; return false }); fmt.Println(count) }", vec!["1"]),
    sync_map_range_empty =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; count := 0; m.Range(func(k, v interface{}) bool { count++; return true }); fmt.Println(count) }", vec!["0"]),
    sync_map_overwrite_existing_key =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); m.Store(\"k\", 2); v, _ := m.Load(\"k\"); fmt.Println(v.(int)) }", vec!["2"]),
    sync_map_swap_replaces_value =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); prev, loaded := m.Swap(\"k\", 2); fmt.Println(prev.(int)); fmt.Println(loaded); v, _ := m.Load(\"k\"); fmt.Println(v.(int)) }", vec!["1", "true", "2"]),
    sync_map_swap_new_key =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; prev, loaded := m.Swap(\"n\", 7); fmt.Println(prev == nil); fmt.Println(loaded); v, _ := m.Load(\"n\"); fmt.Println(v.(int)) }", vec!["true", "false", "7"]),
    sync_map_compare_and_swap_success =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); swapped := m.CompareAndSwap(\"k\", 1, 2); v, _ := m.Load(\"k\"); fmt.Println(swapped); fmt.Println(v.(int)) }", vec!["true", "2"]),
    sync_map_compare_and_swap_fail =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); swapped := m.CompareAndSwap(\"k\", 9, 2); v, _ := m.Load(\"k\"); fmt.Println(swapped); fmt.Println(v.(int)) }", vec!["false", "1"]),
    sync_map_compare_and_delete_success =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); deleted := m.CompareAndDelete(\"k\", 1); _, ok := m.Load(\"k\"); fmt.Println(deleted); fmt.Println(ok) }", vec!["true", "false"]),
    sync_map_compare_and_delete_wrong_old =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); deleted := m.CompareAndDelete(\"k\", 2); _, ok := m.Load(\"k\"); fmt.Println(deleted); fmt.Println(ok) }", vec!["false", "true"]),
    sync_once_do_runs_exactly_once =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var once sync.Once; n := 0; f := func() { n++ }; once.Do(f); once.Do(f); fmt.Println(n) }", vec!["1"]),
    sync_once_do_with_closure_capture =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var once sync.Once; sum := 0; once.Do(func() { sum = 10 }); once.Do(func() { sum = 99 }); fmt.Println(sum) }", vec!["10"]),
    sync_once_separate_instances =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var a sync.Once; var b sync.Once; n := 0; a.Do(func() { n++ }); b.Do(func() { n++ }); fmt.Println(n) }", vec!["2"]),
    sync_once_do_passes_nil_func_panic_guard =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var once sync.Once; ran := false; once.Do(func() { ran = true }); fmt.Println(ran) }", vec!["true"]),
    sync_pool_new_on_empty_get =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var p sync.Pool; p.New = func() interface{} { return 7 }; fmt.Println(p.Get().(int)) }", vec!["7"]),
    sync_pool_put_then_get =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var p sync.Pool; p.Put(9); fmt.Println(p.Get().(int)) }", vec!["9"]),
    sync_pool_get_without_new_returns_nil =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var p sync.Pool; fmt.Println(p.Get() == nil) }", vec!["true"]),
    sync_pool_put_nil_then_get =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var p sync.Pool; p.Put(nil); fmt.Println(p.Get() == nil) }", vec!["true"]),
    sync_pool_struct_new_factory =>
        ("package main; import \"fmt\"; import \"sync\"; type buf struct { data []byte }; func main() { var p sync.Pool; p.New = func() interface{} { return &buf{data: make([]byte, 0, 8)} }; b := p.Get().(*buf); fmt.Println(cap(b.data)) }", vec!["8"]),
    sync_map_multiple_keys_range_count =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"a\", 1); m.Store(\"b\", 2); m.Store(\"c\", 3); n := 0; m.Range(func(k, v interface{}) bool { n++; return true }); fmt.Println(n) }", vec!["3"]),
    sync_map_bool_value =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(true, \"yes\"); v, ok := m.Load(true); fmt.Println(v.(string)); fmt.Println(ok) }", vec!["yes", "true"]),
    sync_map_pointer_key =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { key := new(int); *key = 42; var m sync.Map; m.Store(key, \"ptr\"); v, ok := m.Load(key); fmt.Println(v.(string)); fmt.Println(ok) }", vec!["ptr", "true"]),
    sync_map_load_or_store_on_deleted =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); m.Delete(\"k\"); actual, loaded := m.LoadOrStore(\"k\", 2); fmt.Println(actual.(int)); fmt.Println(loaded) }", vec!["2", "false"]),
    sync_once_nested_do_only_outer_counts =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var outer sync.Once; var inner sync.Once; n := 0; outer.Do(func() { inner.Do(func() { n++ }); inner.Do(func() { n++ }) }); fmt.Println(n) }", vec!["1"]),
    sync_pool_reuse_after_put =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var p sync.Pool; p.New = func() interface{} { return 1 }; first := p.Get().(int); p.Put(first + 10); second := p.Get().(int); fmt.Println(second) }", vec!["11"]),
    sync_map_range_collect_keys =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"x\", 1); m.Store(\"y\", 2); keys := 0; m.Range(func(k, v interface{}) bool { keys++; return true }); fmt.Println(keys) }", vec!["2"]),
    sync_map_zero_value_usable =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(0, 0); v, ok := m.Load(0); fmt.Println(v.(int)); fmt.Println(ok) }", vec!["0", "true"]),
    sync_once_zero_value =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var once sync.Once; n := 0; once.Do(func() { n = 1 }); fmt.Println(n) }", vec!["1"]),
    sync_map_string_to_int_len =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"hello\", len(\"hello\")); v, _ := m.Load(\"hello\"); fmt.Println(v.(int)) }", vec!["5"]),
    sync_map_load_and_delete_then_load_or_store =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); m.LoadAndDelete(\"k\"); actual, loaded := m.LoadOrStore(\"k\", 3); fmt.Println(actual.(int)); fmt.Println(loaded) }", vec!["3", "false"]),
    sync_pool_multiple_put_get_cycle =>
        ("package main; import \"fmt\"; import \"sync\"; func main() { var p sync.Pool; p.Put(1); p.Put(2); a := p.Get().(int); b := p.Get().(int); fmt.Println(a + b) }", vec!["3"]),
    sync_map_nested_struct_in_range =>
        ("package main; import \"fmt\"; import \"sync\"; type pair struct { a int; b int }; func main() { var m sync.Map; m.Store(1, pair{a: 2, b: 3}); sum := 0; m.Range(func(k, v interface{}) bool { p := v.(pair); sum = p.a + p.b; return true }); fmt.Println(sum) }", vec!["5"]),
}

go_compile_cases! {
    sync_map_store_load_compile =>
        "package main; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); _, _ = m.Load(\"k\") }",
    sync_map_load_or_store_compile =>
        "package main; import \"sync\"; func main() { var m sync.Map; _, _ = m.LoadOrStore(\"k\", 1) }",
    sync_map_load_and_delete_compile =>
        "package main; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); _, _ = m.LoadAndDelete(\"k\") }",
    sync_map_range_callback_compile =>
        "package main; import \"sync\"; func main() { var m sync.Map; m.Range(func(k, v interface{}) bool { return true }) }",
    sync_map_swap_compile =>
        "package main; import \"sync\"; func main() { var m sync.Map; _, _ = m.Swap(\"k\", 1) }",
    sync_map_compare_and_swap_compile =>
        "package main; import \"sync\"; func main() { var m sync.Map; _ = m.CompareAndSwap(\"k\", 1, 2) }",
    sync_map_compare_and_delete_compile =>
        "package main; import \"sync\"; func main() { var m sync.Map; _ = m.CompareAndDelete(\"k\", 1) }",
    sync_once_do_compile =>
        "package main; import \"sync\"; func main() { var once sync.Once; once.Do(func() {}) }",
    sync_once_do_with_shared_state_compile =>
        "package main; import \"sync\"; func main() { var once sync.Once; n := 0; once.Do(func() { n++ }); _ = n }",
    sync_pool_get_put_compile =>
        "package main; import \"sync\"; func main() { var p sync.Pool; p.New = func() interface{} { return 0 }; p.Put(p.Get()) }",
    sync_pool_new_assign_compile =>
        "package main; import \"sync\"; func main() { var p sync.Pool; p.New = func() interface{} { return make([]byte, 4) }; _ = p.Get() }",
    sync_cond_wait_compile =>
        "package main; import \"sync\"; func main() { var mu sync.Mutex; cond := sync.NewCond(&mu); cond.Wait() }",
    sync_cond_signal_compile =>
        "package main; import \"sync\"; func main() { var mu sync.Mutex; cond := sync.NewCond(&mu); cond.Signal() }",
    sync_cond_broadcast_compile =>
        "package main; import \"sync\"; func main() { var mu sync.Mutex; cond := sync.NewCond(&mu); cond.Broadcast() }",
    sync_cond_wait_in_loop_compile =>
        "package main; import \"sync\"; func main() { var mu sync.Mutex; cond := sync.NewCond(&mu); ready := false; mu.Lock(); for !ready { cond.Wait() }; mu.Unlock() }",
    sync_cond_new_with_locker_compile =>
        "package main; import \"sync\"; func main() { var rw sync.RWMutex; _ = sync.NewCond(&rw) }",
    sync_pool_put_interface_values_compile =>
        "package main; import \"sync\"; func main() { var p sync.Pool; p.Put(1); p.Put(\"x\"); _ = p.Get() }",
    sync_map_delete_compile =>
        "package main; import \"sync\"; func main() { var m sync.Map; m.Delete(\"k\") }",
    sync_once_in_struct_field_compile =>
        "package main; import \"sync\"; type holder struct { once sync.Once }; func main() { var h holder; h.once.Do(func() {}) }",
    sync_map_in_struct_field_compile =>
        "package main; import \"sync\"; type cache struct { m sync.Map }; func main() { var c cache; c.m.Store(1, 2) }",
}
