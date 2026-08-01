use crate::helpers::run_prints;

#[test]
fn test_thread_created_with_start_false_is_not_alive() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(start = false) {
                println("run")
            }
            println(worker.isAlive)
            println(worker.name)
        }
    "#,
    );
    assert_eq!(out, &["false", "Thread-0"]);
}

#[test]
fn test_thread_start_true_runs_immediately() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread {
                println("ok")
            }
            worker.join()
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_thread_name_is_configurable() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(name = "worker-a", start = false) {}
            println(worker.name)
        }
    "#,
    );
    assert_eq!(out, &["worker-a"]);
}

#[test]
fn test_thread_is_daemon_flag_is_settable_before_start() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(name = "daemon", isDaemon = true, start = false) {}
            println(worker.isDaemon)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_thread_priority_is_settable_before_start() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(name = "prio", priority = Thread.MAX_PRIORITY, start = false) {}
            println(worker.priority)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_thread_join_waits_for_completion() {
    let out = run_prints(
        r#"
        fun main() {
            val marker = java.util.concurrent.atomic.AtomicInteger(0)
            val worker = kotlin.concurrent.thread(start = false) {
                marker.incrementAndGet()
            }
            println(marker.get())
            worker.start()
            worker.join()
            println(marker.get())
            println(worker.isAlive)
        }
    "#,
    );
    assert_eq!(out, &["0", "1", "false"]);
}

#[test]
fn test_thread_cannot_be_started_twice() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(start = false) {}
            worker.start()
            try {
                worker.start()
                println("started-again")
            } catch (ex: Exception) {
                println("error")
            }
        }
    "#,
    );
    assert_eq!(out, &["error"]);
}

#[test]
fn test_thread_exception_handler_captures_background_error() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(start = false, name = "boom") {
                throw RuntimeException("boom")
            }
            worker.setUncaughtExceptionHandler { thread, ex ->
                println(thread.name + ":" + ex.message)
            }
            worker.start()
            worker.join()
        }
    "#,
    );
    assert_eq!(out, &["boom:boom"]);
}

#[test]
fn test_thread_interrupted_before_start_stays_false_after_creation() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(start = false) {}
            println(worker.isInterrupted)
        }
    "#,
    );
    assert_eq!(out, &["false"]);
}

#[test]
fn test_thread_interrupt_flag_set_and_visible() {
    let out = run_prints(
        r#"
        fun main() {
            val latch = java.util.concurrent.CountDownLatch(1)
            var observed = ""
            val worker = kotlin.concurrent.thread(start = false) {
                latch.await()
                if (Thread.currentThread().isInterrupted()) {
                    observed = "interrupted"
                }
            }
            worker.start()
            worker.interrupt()
            latch.countDown()
            worker.join()
            println(observed)
            println(worker.isInterrupted)
        }
    "#,
    );
    assert_eq!(out, &["interrupted", "false"]);
}

#[test]
fn test_thread_interrupted_static_clears_status() {
    let out = run_prints(
        r#"
        fun main() {
            val thread = kotlin.concurrent.thread(start = false) {}
            thread.start()
            thread.interrupt()
            thread.join()
            val before = thread.isInterrupted()
            val fromCurrent = Thread.interrupted()
            println(before)
            println(fromCurrent)
        }
    "#,
    );
    assert_eq!(out, &["false", "false"]);
}

#[test]
fn test_thread_local_state_does_not_leak_between_threads() {
    let out = run_prints(
        r#"
        fun main() {
            val local = java.lang.ThreadLocal<String>()
            local.set("main")
            var childValue = ""
            val worker = kotlin.concurrent.thread {
                childValue = local.get() ?: "unset"
            }
            worker.join()
            println(local.get())
            println(childValue)
        }
    "#,
    );
    assert_eq!(out, &["main", "unset"]);
}

#[test]
fn test_thread_sleep_is_interrupted() {
    let out = run_prints(
        r#"
        fun main() {
            var out = ""
            val worker = kotlin.concurrent.thread(start = false) {
                try {
                    Thread.sleep(10000)
                } catch (ex: InterruptedException) {
                    out = "interrupted"
                }
            }
            worker.start()
            Thread.sleep(10)
            worker.interrupt()
            worker.join()
            println(out)
            println(worker.isAlive)
        }
    "#,
    );
    assert_eq!(out, &["interrupted", "false"]);
}

#[test]
fn test_thread_current_thread_name_from_worker() {
    let out = run_prints(
        r#"
        fun main() {
            var threadName = ""
            val worker = kotlin.concurrent.thread(name = "worker-name") {
                threadName = Thread.currentThread().name
            }
            worker.join()
            println(threadName)
        }
    "#,
    );
    assert_eq!(out, &["worker-name"]);
}

#[test]
fn test_thread_id_is_reported_positive() {
    let out = run_prints(
        r#"
        fun main() {
            var id = 0L
            val worker = kotlin.concurrent.thread {
                id = Thread.currentThread().id
            }
            worker.join()
            println(id > 0)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_thread_state_before_and_after_join() {
    let out = run_prints(
        r#"
        fun main() {
            var before = ""
            val worker = kotlin.concurrent.thread(start = false) {
                Thread.sleep(5)
            }
            before = worker.state.name
            worker.start()
            worker.join()
            println(before)
            println(worker.state.name)
        }
    "#,
    );
    assert_eq!(out, &["NEW", "TERMINATED"]);
}

#[test]
fn test_thread_group_name_for_worker() {
    let out = run_prints(
        r#"
        fun main() {
            var group = ""
            val worker = kotlin.concurrent.thread {
                group = Thread.currentThread().threadGroup.name
            }
            worker.join()
            println(group)
        }
    "#,
    );
    assert_eq!(out, &["main"]);
}

#[test]
fn test_thread_priority_after_start_remains_positive() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(priority = 4, start = false) {}
            worker.start()
            println(worker.priority >= Thread.MIN_PRIORITY)
            worker.join()
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_thread_is_alive_is_true_while_running() {
    let out = run_prints(
        r#"
        fun main() {
            val ready = java.util.concurrent.CountDownLatch(1)
            val canFinish = java.util.concurrent.CountDownLatch(1)
            val running = java.util.concurrent.atomic.AtomicBoolean(false)
            val worker = kotlin.concurrent.thread(start = false) {
                running.set(true)
                ready.countDown()
                canFinish.await()
            }
            worker.start()
            ready.await()
            println(worker.isAlive)
            canFinish.countDown()
            worker.join()
            println(worker.isAlive)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_thread_join_called_after_completion_returns_fast() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread {
                println("done")
            }
            worker.join()
            worker.join()
            println("after")
        }
    "#,
    );
    assert_eq!(out, &["done", "after"]);
}

#[test]
fn test_thread_yield_allows_other_thread_work() {
    let out = run_prints(
        r#"
        fun main() {
            val result = java.util.concurrent.atomic.AtomicInteger(0)
            val worker = kotlin.concurrent.thread {
                var i = 0
                while (i < 3) {
                    result.incrementAndGet()
                    Thread.yield()
                    i += 1
                }
            }
            worker.join()
            println(result.get() == 3)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_thread_interrupt_in_worker_resets_main_flag() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(start = false) {
                try {
                    Thread.sleep(1000)
                    println("complete")
                } catch (ex: InterruptedException) {
                    println("interrupted")
                }
            }
            worker.start()
            worker.interrupt()
            worker.join()
            println(worker.isInterrupted)
            println(worker.isInterrupted())
        }
    "#,
    );
    assert_eq!(out, &["interrupted", "false", "false"]);
}

#[test]
fn test_thread_can_set_uncaught_exception_handler_before_start() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(start = false) {
                throw IllegalStateException("x")
            }
            var observed = "none"
            worker.uncaughtExceptionHandler = java.lang.Thread.UncaughtExceptionHandler { t, e ->
                observed = t.name + ":" + e::class.simpleName!!
            }
            worker.start()
            worker.join()
            println(observed)
        }
    "#,
    );
    assert_eq!(out, &["Thread-0:IllegalStateException"]);
}

#[test]
fn test_thread_can_replace_uncaught_exception_handler() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(start = false) {
                throw Exception("x")
            }
            worker.setUncaughtExceptionHandler { _, _ -> println("first") }
            worker.setUncaughtExceptionHandler { _, _ -> println("second") }
            worker.start()
            worker.join()
        }
    "#,
    );
    assert_eq!(out, &["second"]);
}

#[test]
fn test_thread_with_latch_synchronizes_work() {
    let out = run_prints(
        r#"
        fun main() {
            val start = java.util.concurrent.CountDownLatch(1)
            val ready = java.util.concurrent.CountDownLatch(1)
            val value = java.util.concurrent.atomic.AtomicInteger(0)
            val worker = kotlin.concurrent.thread {
                ready.countDown()
                start.await()
                value.incrementAndGet()
            }
            ready.await()
            println(value.get())
            start.countDown()
            worker.join()
            println(value.get())
        }
    "#,
    );
    assert_eq!(out, &["0", "1"]);
}

#[test]
fn test_thread_sleep_short_and_observes_interrupt() {
    let out = run_prints(
        r#"
        fun main() {
            var tag = ""
            val worker = kotlin.concurrent.thread {
                try {
                    Thread.sleep(20)
                    tag = "slept"
                } catch (ex: InterruptedException) {
                    tag = "interrupted"
                }
            }
            Thread.sleep(5)
            worker.interrupt()
            worker.join()
            println(tag)
        }
    "#,
    );
    assert_eq!(out, &["interrupted"]);
}

#[test]
fn test_thread_uses_default_name_when_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(start = false) {}
            println(worker.name.startsWith("Thread-"))
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_thread_counted_increment_from_many_workers() {
    let out = run_prints(
        r#"
        fun main() {
            val counter = java.util.concurrent.atomic.AtomicInteger(0)
            val done = java.util.concurrent.CountDownLatch(3)
            fun makeWorker() = kotlin.concurrent.thread {
                counter.incrementAndGet()
                done.countDown()
            }
            makeWorker(); makeWorker(); makeWorker()
            done.await()
            println(counter.get())
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_thread_can_run_when_joined_on_multiple_times() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread {
                println("one")
            }
            worker.join()
            worker.join()
            println("done")
        }
    "#,
    );
    assert_eq!(out, &["one", "done"]);
}

#[test]
fn test_thread_is_alive_stays_false_for_unstarted() {
    let out = run_prints(
        r#"
        fun main() {
            val worker = kotlin.concurrent.thread(start = false) {}
            println(worker.isAlive)
        }
    "#,
    );
    assert_eq!(out, &["false"]);
}
