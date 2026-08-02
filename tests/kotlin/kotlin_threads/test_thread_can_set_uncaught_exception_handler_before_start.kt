// vybe-test: kotlin/kotlin_threads/test_thread_can_set_uncaught_exception_handler_before_start
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

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
            __check((observed).toString(), "Thread-0:IllegalStateException")
        }
