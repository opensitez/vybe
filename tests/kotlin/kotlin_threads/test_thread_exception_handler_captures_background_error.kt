// vybe-test: kotlin/kotlin_threads/test_thread_exception_handler_captures_background_error
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread(start = false, name = "boom") {
                throw RuntimeException("boom")
            }
            worker.setUncaughtExceptionHandler { thread, ex ->
                __check((thread.name + ":" + ex.message).toString(), "boom:boom")
            }
            worker.start()
            worker.join()
        }
