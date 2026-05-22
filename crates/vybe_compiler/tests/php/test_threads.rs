use super::helpers::compile_ok;

// ── Thread creation and join ────────────────────────────────
#[test] fn thread_create_join() { compile_ok(r#"<?php
$thread = thread_create(function() {
    return 42;
});
$result = thread_join($thread);
echo $result;
"#); }

#[test] fn thread_with_closure() { compile_ok(r#"<?php
$data = "hello";
$thread = thread_create(function() use ($data) {
    return strtoupper($data);
});
$result = thread_join($thread);
"#); }

#[test] fn multiple_threads() { compile_ok(r#"<?php
$t1 = thread_create(fn() => 1 + 2);
$t2 = thread_create(fn() => 3 + 4);
$r1 = thread_join($t1);
$r2 = thread_join($t2);
echo $r1 + $r2;
"#); }

// ── Mutex ───────────────────────────────────────────────────
#[test] fn mutex_basic() { compile_ok(r#"<?php
$lock = mutex_create();
mutex_lock($lock);
$shared = 42;
mutex_unlock($lock);
"#); }

#[test] fn mutex_in_threads() { compile_ok(r#"<?php
$lock = mutex_create();
$counter = 0;

$t1 = thread_create(function() use ($lock) {
    mutex_lock($lock);
    mutex_unlock($lock);
});

$t2 = thread_create(function() use ($lock) {
    mutex_lock($lock);
    mutex_unlock($lock);
});

thread_join($t1);
thread_join($t2);
"#); }

// ── Fiber + Thread combination ──────────────────────────────
#[test] fn fiber_in_thread() {
    std::thread::Builder::new()
        .stack_size(8 * 1024 * 1024)
        .spawn(|| compile_ok(r#"<?php
$thread = thread_create(function() {
    $fiber = new Fiber(function() {
        Fiber::suspend('from thread fiber');
        return 'done';
    });
    $v = $fiber->start();
    $fiber->resume();
    return $fiber->getReturn();
});
$result = thread_join($thread);
"#))
        .unwrap()
        .join()
        .unwrap();
}
