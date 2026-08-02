// vybe-test: kotlin/result_patterns/test_result_get_or_else_failure_path
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Result.failure<Int>(Exception("boom"))
            __check((value.getOrElse { 99 }).toString(), "99")
        }
