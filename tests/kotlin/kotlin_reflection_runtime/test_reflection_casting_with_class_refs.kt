// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_casting_with_class_refs
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = ProbeImpl(7)
            val cls = ProbeImpl::class
            val casted = cls.java.cast(value)
            __check((casted is ProbeImpl).toString(), "true")
            __check(((casted as ProbeImpl).id).toString(), "7")
        }
