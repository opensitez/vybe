// vybe-test: kotlin/local_functions/test_local_function_name_hides_top_level_function
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun format(value: Int): Int = value

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun format(value: String): String = value + "!"
            __check((format("x")).toString(), "x!")
            __check((format(3)).toString(), "3")
        }
