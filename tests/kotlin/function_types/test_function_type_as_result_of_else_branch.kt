// vybe-test: kotlin/function_types/test_function_type_as_result_of_else_branch
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun pick(upper: Boolean): (Int) -> Int {
            return if (upper) { { it * 2 } } else { { it + 5 } }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(true)(3)).toString(), "6")
            __check((pick(false)(3)).toString(), "8")
        }
