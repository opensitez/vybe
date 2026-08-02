// vybe-test: kotlin/kotlin_lazy_delegates/test_lazy_thread_safety_modes
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

import kotlin.LazyThreadSafetyMode
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var count = 0
            val a by lazy(LazyThreadSafetyMode.NONE) {
                count += 1
                "a"
            }
            val b by lazy(LazyThreadSafetyMode.SYNCHRONIZED) {
                count += 10
                "b"
            }
            __check((a).toString(), "a")
            __check((a).toString(), "a")
            __check((b).toString(), "b")
            __check((b).toString(), "b")
            __check((count).toString(), "11")
        }
