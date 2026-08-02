// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_update_with_lambda
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = java.util.concurrent.atomic.AtomicInteger(1)
            val updated = counter.updateAndGet { value -> value * 3 }
            __check((updated).toString(), "3")
            val finalValue = counter.accumulateAndGet(4) { left, right -> left + right }
            __check((finalValue).toString(), "7")
        }
