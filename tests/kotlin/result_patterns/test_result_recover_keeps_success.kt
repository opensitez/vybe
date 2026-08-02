// vybe-test: kotlin/result_patterns/test_result_recover_keeps_success
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Result.success(2).recover { 11 }
            __check((value.isSuccess).toString(), "true")
            __check((value.getOrNull()).toString(), "2")
        }
