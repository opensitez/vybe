// vybe-test: kotlin/result_patterns/test_result_nested_recover_chain
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { 1 }
                .map { throw Exception("x") }
                .recover { 2 }
                .mapCatching { it + 1 }
            __check((value.getOrNull()).toString(), "3")
        }
