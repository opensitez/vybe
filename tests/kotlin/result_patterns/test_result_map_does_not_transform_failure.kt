// vybe-test: kotlin/result_patterns/test_result_map_does_not_transform_failure
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Result.failure<Int>(Exception("boom")).map { it + 1 }
            __check((value.isSuccess).toString(), "false")
            __check((value.exceptionOrNull()?.message).toString(), "boom")
        }
