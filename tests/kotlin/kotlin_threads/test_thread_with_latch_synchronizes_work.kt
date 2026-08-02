// vybe-test: kotlin/kotlin_threads/test_thread_with_latch_synchronizes_work
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

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
            __check((value.get()).toString(), "0")
            start.countDown()
            worker.join()
            __check((value.get()).toString(), "1")
        }
