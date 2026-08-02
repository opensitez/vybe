// vybe-test: kotlin/result_patterns/test_result_success_with_null_value
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Result.success<String?>(null)
            __check((value.isSuccess).toString(), "true")
            __check((value.getOrNull() == null).toString(), "true")
            __check((value.getOrElse { "fallback" } ?: "none").toString(), "null")
        }
