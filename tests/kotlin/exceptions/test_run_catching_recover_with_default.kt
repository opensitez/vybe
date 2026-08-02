// vybe-test: kotlin/exceptions/test_run_catching_recover_with_default
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bad = runCatching {
                throw Exception("oops")
            }
            __check((bad.isSuccess).toString(), "false")
            __check((bad.isFailure).toString(), "true")
            __check((bad.getOrElse { "fallback" }).toString(), "fallback")
        }
