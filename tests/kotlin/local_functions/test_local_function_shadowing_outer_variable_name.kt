// vybe-test: kotlin/local_functions/test_local_function_shadowing_outer_variable_name
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 2
            fun compute(value: Int): Int = value + 1
            __check((compute(10)).toString(), "11")
            __check((value).toString(), "2")
        }
