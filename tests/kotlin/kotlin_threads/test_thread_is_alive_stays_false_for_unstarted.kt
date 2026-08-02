// vybe-test: kotlin/kotlin_threads/test_thread_is_alive_stays_false_for_unstarted
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread(start = false) {}
            __check((worker.isAlive).toString(), "false")
        }
