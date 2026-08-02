// vybe-test: kotlin/result_patterns/test_result_map_chain_transitions
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { 3 }
                .map { it * 2 }
                .map { it + 1 }
            __check((value.getOrNull()).toString(), "7")
            __check((value.isSuccess).toString(), "true")
        }
