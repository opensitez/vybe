// vybe-test: kotlin/kotlin_result_api/test_result_fold_success_and_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { 2 }
                .fold({ it.toString() }, { "bad" })
            val bad = runCatching<Int> { throw RuntimeException("boom") }
                .fold({ it.toString() }, { e -> e::class.simpleName.toString() })
            __check((ok).toString(), "2")
            __check((bad).toString(), "RuntimeException")
        }
