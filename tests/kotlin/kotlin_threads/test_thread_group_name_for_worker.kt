// vybe-test: kotlin/kotlin_threads/test_thread_group_name_for_worker
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var group = ""
            val worker = kotlin.concurrent.thread {
                group = Thread.currentThread().threadGroup.name
            }
            worker.join()
            __check((group).toString(), "main")
        }
