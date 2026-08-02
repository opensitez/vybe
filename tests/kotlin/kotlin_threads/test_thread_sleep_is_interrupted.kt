// vybe-test: kotlin/kotlin_threads/test_thread_sleep_is_interrupted
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = ""
            val worker = kotlin.concurrent.thread(start = false) {
                try {
                    Thread.sleep(10000)
                } catch (ex: InterruptedException) {
                    out = "interrupted"
                }
            }
            worker.start()
            Thread.sleep(10)
            worker.interrupt()
            worker.join()
            __check((out).toString(), "interrupted")
            __check((worker.isAlive).toString(), "false")
        }
