// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_object_instance_class_equality
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Any = Probe("x")
            val b: Any = Probe("y")
            __check((a::class == b::class).toString(), "true")
            __check((a::class == Probe::class).toString(), "true")
            __check((a::class.java == b::class.java).toString(), "true")
        }
