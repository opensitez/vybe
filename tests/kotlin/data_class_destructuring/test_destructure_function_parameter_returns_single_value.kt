// vybe-test: kotlin/data_class_destructuring/test_destructure_function_parameter_returns_single_value
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class SumPair(val left: Int, val right: Int)

        fun combine(a: Int, b: Int): SumPair = SumPair(a + b, a * b)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (sum, product) = combine(4, 5)
            __check((sum).toString(), "9")
            __check((product).toString(), "20")
        }
