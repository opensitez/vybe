// vybe-test: kotlin/recursion/test_recursion_string_array_joiner
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun join(values: List<String>): String = if (values.isEmpty()) "" else values[0] + if (values.size == 1) "" else "," + join(values.drop(1))
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((join(listOf("a", "b", "c"))).toString(), "a,b,c")
        }
