// vybe-test: kotlin/kotlin_result_api/test_result_exception_message_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching<Int> { 1 / 0 }
                .map { it + 1 }
                .recover { 0 }
            val thrown = runCatching<Int> { 1 / 0 }
                .exceptionOrNull()
            __check((value).toString(), "0")
            __check((thrown?.let { it.message }).toString(), "/ by zero")
        }
