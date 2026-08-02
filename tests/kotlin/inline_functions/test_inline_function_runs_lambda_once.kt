// vybe-test: kotlin/inline_functions/test_inline_function_runs_lambda_once
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun once(block: () -> Int): Int = block()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((once { 5 + 2 }).toString(), "7")
        }
