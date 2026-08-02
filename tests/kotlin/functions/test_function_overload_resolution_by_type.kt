// vybe-test: kotlin/functions/test_function_overload_resolution_by_type
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun format(value: Int): String = "int:" + value
        fun format(value: String): String = "str:" + value

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((format(7)).toString(), "int:7")
            __check((format("7")).toString(), "str:7")
        }
