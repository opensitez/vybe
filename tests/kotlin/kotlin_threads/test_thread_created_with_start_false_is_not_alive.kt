// vybe-test: kotlin/kotlin_threads/test_thread_created_with_start_false_is_not_alive
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun main() {
            val worker = kotlin.concurrent.thread(start = false) {
                println("run")
            }
            println(worker.isAlive)
            println(worker.name)
        }

