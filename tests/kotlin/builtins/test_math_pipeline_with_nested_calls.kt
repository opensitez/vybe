// vybe-test: kotlin/builtins/test_math_pipeline_with_nested_calls
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val score = abs(min(-12, -5) + max(2, 8))
            val amplified = score * score
            __check((score).toString(), "5")
            __check((amplified).toString(), "25")
        }
