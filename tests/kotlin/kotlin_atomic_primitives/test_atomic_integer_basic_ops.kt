// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_integer_basic_ops
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = java.util.concurrent.atomic.AtomicInteger(5)
            __check((counter.get()).toString(), "5")
            __check((counter.incrementAndGet()).toString(), "6")
            __check((counter.getAndAdd(3)).toString(), "6")
            __check((counter.get()).toString(), "9")
        }
