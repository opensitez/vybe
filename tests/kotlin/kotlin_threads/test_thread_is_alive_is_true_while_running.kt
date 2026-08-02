// vybe-test: kotlin/kotlin_threads/test_thread_is_alive_is_true_while_running
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ready = java.util.concurrent.CountDownLatch(1)
            val canFinish = java.util.concurrent.CountDownLatch(1)
            val running = java.util.concurrent.atomic.AtomicBoolean(false)
            val worker = kotlin.concurrent.thread(start = false) {
                running.set(true)
                ready.countDown()
                canFinish.await()
            }
            worker.start()
            ready.await()
            __check((worker.isAlive).toString(), "true")
            canFinish.countDown()
            worker.join()
            __check((worker.isAlive).toString(), "false")
        }
