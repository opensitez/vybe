// vybe-test: kotlin/kotlin_result_api/test_result_is_failure_then_map_not_run
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var seen = 0
            val value = runCatching<Int> { throw Exception("x") }
                .map {
                    seen = 1
                    it
                }
            __check((seen).toString(), "0")
            __check((value.isFailure).toString(), "true")
        }
