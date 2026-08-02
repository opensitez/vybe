// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_reference_compare_and_set_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ref = java.util.concurrent.atomic.AtomicReference(5)
            val ok = ref.compareAndSet(4, 7)
            __check((ok).toString(), "false")
            __check((ref.get()).toString(), "5")
        }
