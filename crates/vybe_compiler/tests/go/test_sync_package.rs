//! sync package: Mutex, RWMutex, WaitGroup, Once, Pool, Map.


go_run_cases! {
    mutex_serial_increment_under_lock => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var mu sync.Mutex; n := 0; for i := 0; i < 5; i++ { mu.Lock(); n++; mu.Unlock() }; fmt.Println(n) }",
        vec!["5"]
    ),
    mutex_trylock_succeeds_when_unlocked => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var mu sync.Mutex; fmt.Println(mu.TryLock()); mu.Unlock() }",
        vec!["true"]
    ),
    mutex_trylock_fails_when_already_locked => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var mu sync.Mutex; mu.Lock(); fmt.Println(mu.TryLock()); mu.Unlock() }",
        vec!["false"]
    ),
    rwmutex_rlock_then_exclusive_lock => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var rw sync.RWMutex; v := 0; rw.RLock(); v = 1; rw.RUnlock(); rw.Lock(); v = 2; rw.Unlock(); fmt.Println(v) }",
        vec!["2"]
    ),
    rwmutex_multiple_rlock_same_goroutine => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var rw sync.RWMutex; rw.RLock(); rw.RLock(); fmt.Println(\"ok\"); rw.RUnlock(); rw.RUnlock() }",
        vec!["ok"]
    ),
    waitgroup_add_done_wait_same_goroutine => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var wg sync.WaitGroup; wg.Add(1); wg.Done(); wg.Wait(); fmt.Println(\"done\") }",
        vec!["done"]
    ),
    waitgroup_zero_value_wait_returns_immediately => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var wg sync.WaitGroup; wg.Wait(); fmt.Println(\"ready\") }",
        vec!["ready"]
    ),
    waitgroup_add_multiple_before_single_done => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var wg sync.WaitGroup; wg.Add(3); wg.Done(); wg.Done(); wg.Done(); wg.Wait(); fmt.Println(0) }",
        vec!["0"]
    ),
    once_do_runs_supplied_function_once => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var once sync.Once; n := 0; f := func() { n++ }; once.Do(f); once.Do(f); fmt.Println(n) }",
        vec!["1"]
    ),
    pool_new_invoked_on_empty_get => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var p sync.Pool; p.New = func() interface{} { return 7 }; fmt.Println(p.Get().(int)) }",
        vec!["7"]
    ),
    pool_put_then_get_returns_stored_value => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var p sync.Pool; p.New = func() interface{} { return 1 }; p.Put(9); fmt.Println(p.Get().(int)) }",
        vec!["9"]
    ),
    pool_get_without_new_returns_nil => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var p sync.Pool; fmt.Println(p.Get() == nil) }",
        vec!["true"]
    ),
    sync_map_store_load_existing_key => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 42); v, ok := m.Load(\"k\"); fmt.Println(v.(int)); fmt.Println(ok) }",
        vec!["42", "true"]
    ),
    sync_map_load_missing_key => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; _, ok := m.Load(\"missing\"); fmt.Println(ok) }",
        vec!["false"]
    ),
    sync_map_load_or_store_keeps_existing => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"a\", 1); actual, loaded := m.LoadOrStore(\"a\", 99); fmt.Println(actual.(int)); fmt.Println(loaded) }",
        vec!["1", "true"]
    ),
    sync_map_load_or_store_inserts_when_absent => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; actual, loaded := m.LoadOrStore(\"b\", 3); fmt.Println(actual.(int)); fmt.Println(loaded) }",
        vec!["3", "false"]
    ),
    sync_map_load_and_delete_removes_entry => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"x\", 5); v, ok := m.LoadAndDelete(\"x\"); _, still := m.Load(\"x\"); fmt.Println(v.(int)); fmt.Println(ok); fmt.Println(still) }",
        vec!["5", "true", "false"]
    ),
    sync_map_delete_removes_key => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"d\", 1); m.Delete(\"d\"); _, ok := m.Load(\"d\"); fmt.Println(ok) }",
        vec!["false"]
    ),
    sync_map_range_accumulates_values => (
        "package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Map; m.Store(\"a\", 10); m.Store(\"b\", 20); sum := 0; m.Range(func(k, v interface{}) bool { sum += v.(int); return true }); fmt.Println(sum) }",
        vec!["30"]
    ),
}

go_compile_cases! {
    mutex_goroutines_increment_shared_counter => "package main; import \"sync\"; func main() { var mu sync.Mutex; n := 0; for i := 0; i < 3; i++ { go func() { mu.Lock(); n++; mu.Unlock() }() }; mu.Lock(); mu.Unlock() }",
    mutex_defer_unlock_on_return => "package main; import \"sync\"; func main() { var mu sync.Mutex; mu.Lock(); defer mu.Unlock(); _ = 1 }",
    rwmutex_concurrent_readers_and_writer => "package main; import \"sync\"; func main() { var rw sync.RWMutex; rw.RLock(); go func() { rw.RLock(); rw.RUnlock() }(); rw.RUnlock() }",
    rwmutex_trylock_after_write_lock => "package main; import \"sync\"; func main() { var rw sync.RWMutex; rw.Lock(); _ = rw.TryRLock(); rw.Unlock() }",
    waitgroup_waits_for_spawned_goroutines => "package main; import \"sync\"; func main() { var wg sync.WaitGroup; wg.Add(1); go func() { defer wg.Done(); _ = 1 }(); wg.Wait() }",
    waitgroup_add_called_from_goroutine => "package main; import \"sync\"; func main() { var wg sync.WaitGroup; ch := make(chan struct{}); go func() { wg.Add(1); close(ch) }(); <-ch; go func() { wg.Done() }(); wg.Wait() }",
    once_concurrent_do_from_goroutines => "package main; import \"sync\"; func main() { var once sync.Once; for i := 0; i < 3; i++ { go once.Do(func() {}) }; once.Do(func() {}) }",
    pool_concurrent_get_put_cycle => "package main; import \"sync\"; func main() { var p sync.Pool; p.New = func() interface{} { return 0 }; go func() { p.Put(p.Get()) }(); _ = p.Get() }",
    sync_map_concurrent_store_and_load => "package main; import \"sync\"; func main() { var m sync.Map; go func() { m.Store(\"k\", 1) }(); _, _ = m.Load(\"k\") }",
    sync_map_swap_replaces_value => "package main; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); _, _ = m.Swap(\"k\", 2) }",
    sync_map_compare_and_swap_updates => "package main; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); _ = m.CompareAndSwap(\"k\", 1, 2) }",
    sync_map_compare_and_delete_removes => "package main; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); _ = m.CompareAndDelete(\"k\", 1) }",
}
