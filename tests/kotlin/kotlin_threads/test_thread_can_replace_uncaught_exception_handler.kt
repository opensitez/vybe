// vybe-test: kotlin/kotlin_threads/test_thread_can_replace_uncaught_exception_handler
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun main() {
            val worker = kotlin.concurrent.thread(start = false) {
                throw Exception("x")
            }
            worker.setUncaughtExceptionHandler { _, _ -> println("first") }
            worker.setUncaughtExceptionHandler { _, _ -> println("second") }
            worker.start()
            worker.join()
        }

