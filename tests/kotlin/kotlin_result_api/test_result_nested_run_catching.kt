// vybe-test: kotlin/kotlin_result_api/test_result_nested_run_catching
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching {
                runCatching { "3".toInt() }.getOrThrow()
            }
            __check((value.isSuccess).toString(), "true")
            __check((value.getOrNull()).toString(), "3")
        }
