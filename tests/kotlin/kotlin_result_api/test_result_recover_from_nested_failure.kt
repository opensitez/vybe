// vybe-test: kotlin/kotlin_result_api/test_result_recover_from_nested_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching {
                runCatching<Int> { "x".toInt() }
            }.recover { Result.failure<Int>(RuntimeException("bad")) }
            __check((value.isSuccess).toString(), "true")
        }
