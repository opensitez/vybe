// vybe-test: kotlin/kotlin_result_api/test_result_fold_failure_message
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching<Int> { 1 / 0 }
                .fold({ "ok" }, { e -> e::class.simpleName.toString() })
            __check((value).toString(), "ArithmeticException")
        }
