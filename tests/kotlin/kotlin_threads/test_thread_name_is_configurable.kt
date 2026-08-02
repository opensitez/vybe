// vybe-test: kotlin/kotlin_threads/test_thread_name_is_configurable
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread(name = "worker-a", start = false) {}
            __check((worker.name).toString(), "worker-a")
        }
