// vybe-test: kotlin/kotlin_threads/test_thread_state_before_and_after_join
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var before = ""
            val worker = kotlin.concurrent.thread(start = false) {
                Thread.sleep(5)
            }
            before = worker.state.name
            worker.start()
            worker.join()
            __check((before).toString(), "NEW")
            __check((worker.state.name).toString(), "TERMINATED")
        }
