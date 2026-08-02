// vybe-test: kotlin/kotlin_threads/test_thread_priority_after_start_remains_positive
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread(priority = 4, start = false) {}
            worker.start()
            __check((worker.priority >= Thread.MIN_PRIORITY).toString(), "true")
            worker.join()
        }
