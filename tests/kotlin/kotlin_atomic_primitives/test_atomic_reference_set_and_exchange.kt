// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_reference_set_and_exchange
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ref = java.util.concurrent.atomic.AtomicReference("a")
            __check((ref.get()).toString(), "a")
            __check((ref.getAndSet("b")).toString(), "a")
            __check((ref.compareAndSet("b", "c")).toString(), "true")
            __check((ref.get()).toString(), "c")
        }
