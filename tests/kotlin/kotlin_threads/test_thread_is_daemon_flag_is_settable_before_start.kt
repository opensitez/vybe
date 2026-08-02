// vybe-test: kotlin/kotlin_threads/test_thread_is_daemon_flag_is_settable_before_start
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread(name = "daemon", isDaemon = true, start = false) {}
            __check((worker.isDaemon).toString(), "true")
        }
