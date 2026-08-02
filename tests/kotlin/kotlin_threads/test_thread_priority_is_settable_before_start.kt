// vybe-test: kotlin/kotlin_threads/test_thread_priority_is_settable_before_start
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread(name = "prio", priority = Thread.MAX_PRIORITY, start = false) {}
            __check((worker.priority).toString(), "10")
        }
