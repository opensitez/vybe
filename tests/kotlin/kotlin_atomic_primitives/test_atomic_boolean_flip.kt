// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_boolean_flip
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val flag = java.util.concurrent.atomic.AtomicBoolean(false)
            __check((flag.compareAndSet(false, true)).toString(), "true")
            __check((flag.getAndSet(false)).toString(), "true")
            __check((flag.get()).toString(), "false")
            __check((flag.compareAndSet(true, false)).toString(), "false")
        }
