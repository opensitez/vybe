// vybe-test: kotlin/functions/test_function_call_chain_with_hof
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun transform(x: Int, op: (Int) -> Int): Int {
            return op(x)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = transform(10) { it + 5 }
            __check((result).toString(), "15")
        }
