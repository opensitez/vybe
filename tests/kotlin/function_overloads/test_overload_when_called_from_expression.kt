// vybe-test: kotlin/function_overloads/test_overload_when_called_from_expression
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun calc(v: Int): Int = v * 2
        fun calc(v: String): String = v + "!"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((calc(4) + 1).toString(), "9")
            __check((calc("x")).toString(), "x!")
        }
