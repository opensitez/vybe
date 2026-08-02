// vybe-test: kotlin/kotlin_result_api/test_result_fold_return_type_change
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = runCatching { 9 }
            val out = a.fold(
                onSuccess = { v -> v + 1 },
                onFailure = { _ -> 0 }
            )
            __check((out).toString(), "10")
        }
