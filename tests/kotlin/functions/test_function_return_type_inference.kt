// vybe-test: kotlin/functions/test_function_return_type_inference
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun compute(value: Int) = if (value > 0) value.toString() else value

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((compute(4)).toString(), "4")
            __check((compute(-1)).toString(), "-1")
        }
