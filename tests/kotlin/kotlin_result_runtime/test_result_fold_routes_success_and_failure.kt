// vybe-test: kotlin/kotlin_result_runtime/test_result_fold_routes_success_and_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { "7".toInt() }
                .fold(
                    onSuccess = { value -> "ok-" + value.toString() },
                    onFailure = { "bad" }
                )
            val bad = runCatching<Int> { "x".toInt() }
                .fold(
                    onSuccess = { value -> "ok-" + value.toString() },
                    onFailure = { e -> "err-" + e::class.simpleName.toString() }
                )
            __check((ok).toString(), "ok-7")
            __check((bad).toString(), "err-NumberFormatException")
        }
