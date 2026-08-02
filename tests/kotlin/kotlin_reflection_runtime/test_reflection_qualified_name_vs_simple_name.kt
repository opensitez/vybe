// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_qualified_name_vs_simple_name
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = MarkerContract::class
            __check((c.qualifiedName?.contains("MarkerContract")).toString(), "true")
            __check((c.simpleName).toString(), "MarkerContract")
        }
