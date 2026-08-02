// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_integer_compare_and_set
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = java.util.concurrent.atomic.AtomicInteger(10)
            val first = counter.compareAndSet(10, 20)
            val second = counter.compareAndSet(10, 30)
            __check((first).toString(), "true")
            __check((second).toString(), "false")
            __check((counter.get()).toString(), "20")
        }
