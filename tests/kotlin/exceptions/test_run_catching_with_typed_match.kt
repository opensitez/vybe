// vybe-test: kotlin/exceptions/test_run_catching_with_typed_match
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = runCatching {
                throw IllegalArgumentException("bad")
            }

            val message = result.exceptionOrNull()?.let { it.message } ?: "none"
            __check((result.isFailure).toString(), "true")
            __check((message).toString(), "bad")
        }
