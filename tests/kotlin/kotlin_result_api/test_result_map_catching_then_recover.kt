// vybe-test: kotlin/kotlin_result_api/test_result_map_catching_then_recover
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { 1 }
                .mapCatching { throw IllegalArgumentException("nope") }
                .recover { 99 }
            __check((value.getOrNull()).toString(), "99")
        }
