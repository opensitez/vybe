// vybe-test: kotlin/kotlin_result_api/test_result_mixed_chain_map_then_recover
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching<Int> { "x".toInt() }
                .map { it * 2 }
                .recover { 21 }
            __check((value.getOrElse { 0 }).toString(), "21")
        }
