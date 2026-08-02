// vybe-test: kotlin/kotlin_threads/test_thread_current_thread_name_from_worker
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var threadName = ""
            val worker = kotlin.concurrent.thread(name = "worker-name") {
                threadName = Thread.currentThread().name
            }
            worker.join()
            __check((threadName).toString(), "worker-name")
        }
