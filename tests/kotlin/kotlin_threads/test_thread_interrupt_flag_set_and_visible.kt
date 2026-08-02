// vybe-test: kotlin/kotlin_threads/test_thread_interrupt_flag_set_and_visible
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val latch = java.util.concurrent.CountDownLatch(1)
            var observed = ""
            val worker = kotlin.concurrent.thread(start = false) {
                latch.await()
                if (Thread.currentThread().isInterrupted()) {
                    observed = "interrupted"
                }
            }
            worker.start()
            worker.interrupt()
            latch.countDown()
            worker.join()
            __check((observed).toString(), "interrupted")
            __check((worker.isInterrupted).toString(), "false")
        }
