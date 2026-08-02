// vybe-test: kotlin/kotlin_threads/test_thread_join_called_after_completion_returns_fast
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread {
                __check(("done").toString(), "done")
            }
            worker.join()
            worker.join()
            __check(("after").toString(), "after")
        }
