// vybe-test: kotlin/result_patterns/test_result_chain_with_recover_and_map
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { throw Exception("x") }
                .recover { 4 }
                .map { it + 1 }
            __check((value.getOrNull()).toString(), "5")
        }
