// vybe-test: kotlin/kotlin_threads/test_thread_interrupted_static_clears_status
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val thread = kotlin.concurrent.thread(start = false) {}
            thread.start()
            thread.interrupt()
            thread.join()
            val before = thread.isInterrupted()
            val fromCurrent = Thread.interrupted()
            __check((before).toString(), "false")
            __check((fromCurrent).toString(), "false")
        }
