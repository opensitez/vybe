// vybe-test: kotlin/result_patterns/test_result_recover_catching_with_success
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Result.failure<Int>(Exception("bad")).recoverCatching { throw Exception("oops") }
            __check((value.isSuccess).toString(), "false")
            __check((value.exceptionOrNull()?.message).toString(), "oops")
        }
