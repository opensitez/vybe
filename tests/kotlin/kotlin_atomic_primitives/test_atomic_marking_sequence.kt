// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_marking_sequence
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

import java.util.concurrent.atomic.AtomicInteger

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = AtomicInteger(1)
            var state = ""
            repeat(3) {
                state += a.getAndIncrement().toString() + ","
            }
            __check((state).toString(), "1,2,3,")
            __check((a.get()).toString(), "4")
        }
