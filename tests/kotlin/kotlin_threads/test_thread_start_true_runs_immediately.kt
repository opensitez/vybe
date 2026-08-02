// vybe-test: kotlin/kotlin_threads/test_thread_start_true_runs_immediately
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread {
                __check(("ok").toString(), "ok")
            }
            worker.join()
        }
