// vybe-test: kotlin/result_patterns/test_result_get_or_default_success_and_failure
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val success = Result.success(2)
            val fail = Result.failure<Int>(Exception("boom"))
            __check((success.getOrDefault(9)).toString(), "2")
            __check((fail.getOrDefault(9)).toString(), "9")
        }
