// vybe-test: kotlin/exceptions/test_run_catching_map_and_recover_flow
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { "k".toInt() }
                .map { it + 1 }
                .onFailure { __check(("fail").toString(), "fail") }
                .recover { 9 }

            __check((value.getOrNull()).toString(), "9")
            __check((value.isFailure).toString(), "false")
        }
