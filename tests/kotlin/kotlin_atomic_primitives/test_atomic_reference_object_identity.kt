// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_reference_object_identity
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            data class PairBox(val value: Int)
            val a = PairBox(1)
            val b = PairBox(1)
            val ref = java.util.concurrent.atomic.AtomicReference(a)
            __check((ref.compareAndSet(a, b)).toString(), "true")
            __check((ref.get().value).toString(), "1")
        }
