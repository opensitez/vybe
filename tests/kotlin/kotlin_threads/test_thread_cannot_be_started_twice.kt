// vybe-test: kotlin/kotlin_threads/test_thread_cannot_be_started_twice
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun main() {
            val worker = kotlin.concurrent.thread(start = false) {}
            worker.start()
            try {
                worker.start()
                println("started-again")
            } catch (ex: Exception) {
                println("error")
            }
        }

