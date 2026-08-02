// vybe-test: kotlin/exceptions/test_run_catching_ok_path_preserves_value
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val good = runCatching {
                val value = 3 + 4
                value * 2
            }
            __check((good.isSuccess).toString(), "true")
            __check((good.getOrNull()).toString(), "14")
            __check((good.getOrElse { 0 }).toString(), "14")
        }
