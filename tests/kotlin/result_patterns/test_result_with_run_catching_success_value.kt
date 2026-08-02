// vybe-test: kotlin/result_patterns/test_result_with_run_catching_success_value
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { "x" + "y" }
            __check((value.getOrNull()).toString(), "xy")
        }
