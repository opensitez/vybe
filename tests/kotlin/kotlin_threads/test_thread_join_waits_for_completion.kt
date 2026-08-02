// vybe-test: kotlin/kotlin_threads/test_thread_join_waits_for_completion
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val marker = java.util.concurrent.atomic.AtomicInteger(0)
            val worker = kotlin.concurrent.thread(start = false) {
                marker.incrementAndGet()
            }
            __check((marker.get()).toString(), "0")
            worker.start()
            worker.join()
            __check((marker.get()).toString(), "1")
            __check((worker.isAlive).toString(), "false")
        }
