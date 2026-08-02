// vybe-test: kotlin/kotlin_threads/test_thread_interrupt_in_worker_resets_main_flag
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun main() {
            val worker = kotlin.concurrent.thread(start = false) {
                try {
                    Thread.sleep(1000)
                    println("complete")
                } catch (ex: InterruptedException) {
                    println("interrupted")
                }
            }
            worker.start()
            worker.interrupt()
            worker.join()
            println(worker.isInterrupted)
            println(worker.isInterrupted())
        }

