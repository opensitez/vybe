// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_class_simple_name
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Probe("a")
            __check((value::class.simpleName).toString(), "Probe")
            __check((Probe::class.simpleName).toString(), "Probe")
        }
