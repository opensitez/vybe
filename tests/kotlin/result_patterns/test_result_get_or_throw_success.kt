// vybe-test: kotlin/result_patterns/test_result_get_or_throw_success
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { 21 }
            __check((value.getOrThrow()).toString(), "21")
        }
