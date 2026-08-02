// vybe-test: kotlin/functions/test_function_pass_function_as_value
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun apply(value: Int, op: (Int) -> Int): Int {
            return op(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val transform = { x: Int -> x * 3 }
            __check((apply(4, transform)).toString(), "12")
        }
