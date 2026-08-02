// vybe-test: kotlin/function_types/test_function_type_with_default_argument
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun run(value: Int, op: (Int) -> Int = { it + 1 }): Int = op(value)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((run(2)).toString(), "3")
            __check((run(2, { it * 3 })).toString(), "6")
        }
