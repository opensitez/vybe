//! Channel, select, and concurrency language semantics — one rule per test.

go_run_cases! {
    chan_make_unbuffered_zero_cap => ("package main; import \"fmt\"; func main() { ch := make(chan int); fmt.Println(cap(ch)) }", vec!["0"]),
    chan_buffered_send_recv => ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 9; fmt.Println(<-ch) }", vec!["9"]),
    chan_close_then_zero_value => ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 1; close(ch); v, ok := <-ch; fmt.Println(v, ok) }", vec!["1 false"]),
    chan_range_after_close => ("package main; import \"fmt\"; func main() { ch := make(chan int, 2); ch <- 1; ch <- 2; close(ch); n := 0; for range ch { n++ }; fmt.Println(n) }", vec!["2"]),
    select_receive_ready => ("package main; import \"fmt\"; func main() { ch := make(chan int, 1); ch <- 5; select { case v := <-ch: fmt.Println(v) } }", vec!["5"]),
    select_default_nonblocking => ("package main; import \"fmt\"; func main() { ch := make(chan int); select { case <-ch: fmt.Println(\"recv\"); default: fmt.Println(\"def\") } }", vec!["def"]),
    select_multiple_case_first_wins => ("package main; import \"fmt\"; func main() { a := make(chan int, 1); b := make(chan int, 1); a <- 1; b <- 2; select { case v := <-a: fmt.Println(v); case v := <-b: fmt.Println(v) } }", vec!["1"]),
    sync_mutex_lock_unlock => ("package main; import \"fmt\"; import \"sync\"; func main() { var m sync.Mutex; m.Lock(); m.Unlock(); fmt.Println(\"ok\") }", vec!["ok"]),
    sync_waitgroup_add_done => ("package main; import \"fmt\"; import \"sync\"; func main() { var wg sync.WaitGroup; wg.Add(1); wg.Done(); wg.Wait(); fmt.Println(1) }", vec!["1"]),
    sync_once_do => ("package main; import \"fmt\"; import \"sync\"; func main() { var o sync.Once; n := 0; o.Do(func() { n++ }); o.Do(func() { n++ }); fmt.Println(n) }", vec!["1"]),
}

go_compile_cases! {
    chan_direction_assign_compat => "package main; func main() { ch := make(chan int, 1); var send chan<- int = ch; var recv <-chan int = ch; _ = send; _ = recv }",
    select_send_case_compile => "package main; func main() { ch := make(chan int, 1); select { case ch <- 1: } }",
    select_assign_recv_compile => "package main; func main() { ch := make(chan int, 1); ch <- 1; var v int; select { case v = <-ch: _ = v } }",
    go_closure_spawn_compile => "package main; func main() { go func() {}() }",
    double_close_compile => "package main; func main() { ch := make(chan int); close(ch); close(ch) }",
    send_on_closed_chan_compile => "package main; func main() { ch := make(chan int); close(ch); ch <- 1 }",
    sync_map_store_load => "package main; import \"sync\"; func main() { var m sync.Map; m.Store(\"k\", 1); _, _ = m.Load(\"k\") }",
    sync_pool_put_get => "package main; import \"sync\"; func main() { p := sync.Pool{New: func() interface{} { return 0 }}; p.Put(1); _ = p.Get() }",
    sync_rwlock => "package main; import \"sync\"; func main() { var rw sync.RWMutex; rw.RLock(); rw.RUnlock() }",
    context_with_cancel => "package main; import \"context\"; func main() { _, cancel := context.WithCancel(context.Background()); cancel() }",
}
