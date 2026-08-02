// vybe-test: kotlin/kotlin_result_api/test_result_get_or_null_for_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bad = runCatching<Int> { "x".toInt() }
            __check((bad.getOrNull() == null).toString(), "true")
            __check((bad.exceptionOrNull()?.let { it::class.simpleName }).toString(), "NumberFormatException")
        }
