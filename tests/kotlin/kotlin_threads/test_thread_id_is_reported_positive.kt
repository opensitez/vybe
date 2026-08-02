// vybe-test: kotlin/kotlin_threads/test_thread_id_is_reported_positive
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var id = 0L
            val worker = kotlin.concurrent.thread {
                id = Thread.currentThread().id
            }
            worker.join()
            __check((id > 0).toString(), "true")
        }
