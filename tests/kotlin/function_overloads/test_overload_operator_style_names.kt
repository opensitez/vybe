// vybe-test: kotlin/function_overloads/test_overload_operator_style_names
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun plus(a: Int): Int = a
        fun plus(a: String): String = a
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((plus(5)).toString(), "5")
            __check((plus("y")).toString(), "y")
        }
