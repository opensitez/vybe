// vybe-test: kotlin/kotlin_threads/test_thread_yield_allows_other_thread_work
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun main() {
            val result = java.util.concurrent.atomic.AtomicInteger(0)
            val worker = kotlin.concurrent.thread {
                var i = 0
                while (i < 3) {
                    result.incrementAndGet()
                    Thread.yield()
                    i += 1
                }
            }
            worker.join()
            println(result.get() == 3)
        }

