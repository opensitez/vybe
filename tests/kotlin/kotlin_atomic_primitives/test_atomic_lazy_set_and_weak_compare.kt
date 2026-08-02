// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_lazy_set_and_weak_compare
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.util.concurrent.atomic.AtomicInteger(0)
            value.lazySet(9)
            __check((value.get()).toString(), "9")
            val ok = value.compareAndSet(9, 10)
            __check((ok).toString(), "true")
            __check((value.get()).toString(), "10")
        }
