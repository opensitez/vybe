//! runtime, runtime/debug, runtime/metrics, runtime/pprof, runtime/trace,
//! testing, testing/fstest, testing/iotest, testing/quick - one API per compile smoke.


go_compile_cases! {
    // runtime - core introspection
    runtime_num_cpu => "package main; import \"runtime\"; func main() { _ = runtime.NumCPU() }",
    runtime_func_for_pc => "package main; import \"runtime\"; func main() { _ = runtime.FuncForPC(0) }",
    runtime_set_finalizer => "package main; import \"runtime\"; type T struct{}; func main() { var t T; runtime.SetFinalizer(&t, func(*T) {}) }",
    runtime_gc => "package main; import \"runtime\"; func main() { runtime.GC() }",
    runtime_goos => "package main; import \"runtime\"; func main() { _ = runtime.GOOS }",
    runtime_goarch => "package main; import \"runtime\"; func main() { _ = runtime.GOARCH }",
    runtime_goroutine_profile => "package main; import \"runtime\"; func main() { _, _ = runtime.GoroutineProfile(nil) }",
    runtime_stack => "package main; import \"runtime\"; func main() { buf := make([]byte, 64); _ = runtime.Stack(buf, false) }",
    runtime_num_cgo_call => "package main; import \"runtime\"; func main() { _ = runtime.NumCgoCall() }",
    runtime_cgo_enabled => "package main; import \"runtime\"; func main() { _ = runtime.CgoEnabled }",
    runtime_set_mem_profile_rate => "package main; import \"runtime\"; func main() { _ = runtime.SetMemProfileRate(1000) }",
    runtime_breakpoint => "package main; import \"runtime\"; func main() { runtime.Breakpoint() }",
    runtime_set_block_profile_rate => "package main; import \"runtime\"; func main() { runtime.SetBlockProfileRate(1) }",
    runtime_set_mutex_profile_fraction => "package main; import \"runtime\"; func main() { runtime.SetMutexProfileFraction(1) }",
    runtime_nanotime => "package main; import \"runtime\"; func main() { _ = runtime.Nanotime() }",
    runtime_lock_os_thread => "package main; import \"runtime\"; func main() { runtime.LockOSThread() }",
    runtime_unlock_os_thread => "package main; import \"runtime\"; func main() { runtime.UnlockOSThread() }",
    runtime_caller => "package main; import \"runtime\"; func main() { _, _, _, _ = runtime.Caller(0) }",
    runtime_callers => "package main; import \"runtime\"; func main() { pcs := make([]uintptr, 8); _, _ = runtime.Callers(0, pcs) }",
    runtime_callers_frames => "package main; import \"runtime\"; func main() { pcs := make([]uintptr, 4); n := runtime.Callers(0, pcs); frames := runtime.CallersFrames(pcs[:n]); _, _ = frames.Next() }",
    runtime_num_goroutine => "package main; import \"runtime\"; func main() { _ = runtime.NumGoroutine() }",
    runtime_gomaxprocs => "package main; import \"runtime\"; func main() { _ = runtime.GOMAXPROCS(0) }",
    runtime_keep_alive => "package main; import \"runtime\"; func main() { x := 1; runtime.KeepAlive(x) }",
    runtime_set_cpu_profile_rate => "package main; import \"runtime\"; func main() { runtime.SetCPUProfileRate(100) }",
    runtime_free_os_memory => "package main; import \"runtime\"; func main() { runtime.FreeOSMemory() }",
    runtime_version => "package main; import \"runtime\"; func main() { _ = runtime.Version() }",

    // runtime/debug
    debug_write_heap_dump => "package main; import \"runtime/debug\"; import \"os\"; func main() { debug.WriteHeapDump(os.Stdout.Fd()) }",
    debug_read_build_info => "package main; import \"runtime/debug\"; func main() { _, _ = debug.ReadBuildInfo() }",
    debug_print_stack => "package main; import \"runtime/debug\"; func main() { debug.PrintStack() }",
    debug_set_gc_percent => "package main; import \"runtime/debug\"; func main() { _ = debug.SetGCPercent(100) }",
    debug_set_memory_limit => "package main; import \"runtime/debug\"; func main() { _ = debug.SetMemoryLimit(-1) }",
    debug_set_max_threads => "package main; import \"runtime/debug\"; func main() { _ = debug.SetMaxThreads(10000) }",
    debug_stack => "package main; import \"runtime/debug\"; func main() { _ = debug.Stack() }",

    // runtime/metrics
    metrics_all => "package main; import \"runtime/metrics\"; func main() { _ = metrics.All() }",
    metrics_description => "package main; import \"runtime/metrics\"; func main() { desc := metrics.All(); if len(desc) > 0 { _ = desc[0].Name } }",

    // runtime/pprof
    pprof_lookup => "package main; import \"runtime/pprof\"; func main() { _ = pprof.Lookup(\"goroutine\") }",
    pprof_profiles => "package main; import \"runtime/pprof\"; func main() { _ = pprof.Profiles() }",
    pprof_start_cpu_profile => "package main; import \"runtime/pprof\"; import \"os\"; func main() { _ = pprof.StartCPUProfile(os.Stdout) }",
    pprof_stop_cpu_profile => "package main; import \"runtime/pprof\"; func main() { pprof.StopCPUProfile() }",
    pprof_write_heap_profile => "package main; import \"runtime/pprof\"; import \"os\"; func main() { _ = pprof.WriteHeapProfile(os.Stdout) }",
    pprof_set_goroutine_labels => "package main; import \"runtime/pprof\"; func main() { pprof.SetGoroutineLabels(nil) }",
    pprof_new_profile => "package main; import \"runtime/pprof\"; func main() { _ = pprof.NewProfile(\"custom\") }",

    // runtime/trace
    trace_start => "package main; import \"runtime/trace\"; import \"os\"; func main() { _ = trace.Start(os.Stdout) }",
    trace_stop => "package main; import \"runtime/trace\"; func main() { trace.Stop() }",
    trace_is_enabled => "package main; import \"runtime/trace\"; func main() { _ = trace.IsEnabled() }",
    trace_with_region => "package main; import \"runtime/trace\"; import \"context\"; func main() { ctx, task := trace.NewTask(context.Background(), \"job\"); defer task.End(); _ = trace.WithRegion(ctx, \"step\", func() {}) }",
    trace_log => "package main; import \"runtime/trace\"; import \"context\"; func main() { trace.Log(context.Background(), \"key\", \"val\") }",
    trace_logf => "package main; import \"runtime/trace\"; import \"context\"; func main() { trace.Logf(context.Background(), \"n=%d\", 1) }",

    // testing
    testing_allocs_per_run => "package main; import \"testing\"; func main() { _ = testing.AllocsPerRun(1, func() { _ = make([]byte, 8) }) }",
    testing_coverage => "package main; import \"testing\"; func main() { _ = testing.Coverage() }",
    testing_cover_mode => "package main; import \"testing\"; func main() { _ = testing.CoverMode() }",
    testing_short => "package main; import \"testing\"; func main() { _ = testing.Short() }",
    testing_verbose => "package main; import \"testing\"; func main() { _ = testing.Verbose() }",
    testing_main_start => "package main; import \"testing\"; func main() { m := testing.MainStart(nil, nil, nil, nil); _ = m }",

    // testing/fstest
    fstest_map_fs_empty => "package main; import \"testing/fstest\"; func main() { m := fstest.MapFS{}; _ = len(m) }",
    fstest_map_file_empty => "package main; import \"testing/fstest\"; func main() { f := fstest.MapFile{}; _ = len(f.Data) }",
    fstest_test_os_dir => "package main; import \"testing/fstest\"; import \"os\"; func main() { _ = fstest.TestFS(os.DirFS(\".\"), \"Cargo.toml\") }",
    fstest_map_file_name => "package main; import \"testing/fstest\"; func main() { f := fstest.MapFile{}; _ = f.Name() }",
    fstest_map_file_sys => "package main; import \"testing/fstest\"; func main() { f := fstest.MapFile{}; _ = f.Sys() }",

    // testing/iotest
    iotest_data_err_reader => "package main; import \"testing/iotest\"; import \"strings\"; func main() { _ = iotest.DataErrReader(strings.NewReader(\"abc\")) }",
    iotest_one_byte_reader => "package main; import \"testing/iotest\"; import \"strings\"; func main() { _ = iotest.OneByteReader(strings.NewReader(\"abc\")) }",
    iotest_timeout_reader => "package main; import \"testing/iotest\"; import \"strings\"; func main() { _ = iotest.TimeoutReader(strings.NewReader(\"abc\")) }",
    iotest_truncate_writer => "package main; import \"testing/iotest\"; import \"bytes\"; func main() { _ = iotest.TruncateWriter(bytes.NewBuffer(nil), 4) }",
    iotest_new_read_logger => "package main; import \"testing/iotest\"; import \"strings\"; func main() { _ = iotest.NewReadLogger(strings.NewReader(\"x\")) }",
    iotest_new_write_logger => "package main; import \"testing/iotest\"; import \"bytes\"; func main() { _ = iotest.NewWriteLogger(bytes.NewBuffer(nil)) }",
    iotest_err_reader => "package main; import \"testing/iotest\"; func main() { _ = iotest.ErrReader }",
    iotest_half_reader => "package main; import \"testing/iotest\"; import \"strings\"; func main() { _ = iotest.HalfReader(strings.NewReader(\"abcd\")) }",

    // testing/quick
    quick_check => "package main; import \"testing/quick\"; func main() { f := func(x int) bool { return true }; _ = quick.Check(f, nil) }",
    quick_check_equal => "package main; import \"testing/quick\"; func main() { f := func(x int) bool { return true }; _ = quick.CheckEqual(f, f, nil) }",
    quick_value => "package main; import \"testing/quick\"; func main() { var x int; _ = quick.Value(&x, nil) }",
    quick_config => "package main; import \"testing/quick\"; func main() { cfg := &quick.Config{MaxCount: 10}; f := func(x int) bool { return true }; _ = quick.Check(f, cfg) }",
}
