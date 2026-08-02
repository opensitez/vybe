// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_generic_type_reference
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val probe = Probe("abc")
            val values: List<Any> = listOf(probe, 1, "x")
            __check((values.map { it::class.simpleName }.joinToString(",")).toString(), "Probe,Int,String")
            val first = values[0]::class
            __check((first == Probe::class).toString(), "true")
        }
