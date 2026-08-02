// vybe-test: kotlin/functions/test_function_higher_order_with_default_lambda
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun apply(value: Int, op: (Int) -> Int = { it + 1 }): Int {
            return op(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((apply(4)).toString(), "5")
            __check((apply(4) { it * 3 }).toString(), "12")
        }
