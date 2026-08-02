// vybe-test: kotlin/function_types/test_function_type_with_trailing_lambda_call_style
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun use(v: Int, block: (Int) -> Int): Int = block(v)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((use(4) { it + 10 }).toString(), "14")
        }
