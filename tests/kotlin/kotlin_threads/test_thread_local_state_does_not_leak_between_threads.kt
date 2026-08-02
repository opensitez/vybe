// vybe-test: kotlin/kotlin_threads/test_thread_local_state_does_not_leak_between_threads
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val local = java.lang.ThreadLocal<String>()
            local.set("main")
            var childValue = ""
            val worker = kotlin.concurrent.thread {
                childValue = local.get() ?: "unset"
            }
            worker.join()
            __check((local.get()).toString(), "main")
            __check((childValue).toString(), "unset")
        }
