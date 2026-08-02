// vybe-test: kotlin/kotlin_threads/test_thread_can_run_when_joined_on_multiple_times
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread {
                __check(("one").toString(), "one")
            }
            worker.join()
            worker.join()
            __check(("done").toString(), "done")
        }
