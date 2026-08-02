// vybe-test: kotlin/kotlin_threads/test_thread_counted_increment_from_many_workers
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = java.util.concurrent.atomic.AtomicInteger(0)
            val done = java.util.concurrent.CountDownLatch(3)
            fun makeWorker() = kotlin.concurrent.thread {
                counter.incrementAndGet()
                done.countDown()
            }
            makeWorker()
makeWorker()
makeWorker()
            done.await()
            __check((counter.get()).toString(), "3")
        }
