// vybe-test: kotlin/kotlin_result_api/test_result_nested_failure_bubble
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching {
                runCatching<Int> { "x".toInt() }.getOrThrow()
            }
            __check((value.isFailure).toString(), "true")
            __check((value.exceptionOrNull()?.let { it::class.simpleName }).toString(), "NumberFormatException")
        }
