// vybe-test: kotlin/kotlin_result_api/test_result_map_catching_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bad = runCatching { "x".toInt() }
                .mapCatching { throw IllegalStateException("mapped") }
            __check((bad.getOrNull() == null).toString(), "true")
            __check((bad.exceptionOrNull()?.message).toString(), "mapped")
        }
