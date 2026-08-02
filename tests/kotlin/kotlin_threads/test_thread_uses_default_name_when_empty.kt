// vybe-test: kotlin/kotlin_threads/test_thread_uses_default_name_when_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val worker = kotlin.concurrent.thread(start = false) {}
            __check((worker.name.startsWith("Thread-")).toString(), "true")
        }
