// vybe-test: kotlin/kotlin_threads/test_thread_sleep_short_and_observes_interrupt
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var tag = ""
            val worker = kotlin.concurrent.thread {
                try {
                    Thread.sleep(20)
                    tag = "slept"
                } catch (ex: InterruptedException) {
                    tag = "interrupted"
                }
            }
            Thread.sleep(5)
            worker.interrupt()
            worker.join()
            __check((tag).toString(), "interrupted")
        }
