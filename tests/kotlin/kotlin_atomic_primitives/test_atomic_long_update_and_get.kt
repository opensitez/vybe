// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_long_update_and_get
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.util.concurrent.atomic.AtomicLong(0L)
            value.addAndGet(15)
            __check((value.incrementAndGet()).toString(), "16")
            __check((value.getAndSet(100)).toString(), "16")
            __check((value.get()).toString(), "100")
        }
