// vybe-test: kotlin/kotlin_result_api/test_result_on_success_side_effect
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var marker = ""
            runCatching { 4 }
                .onSuccess { marker += "ok" + it.toString() }
                .onFailure { marker += "fail" }
            __check((marker).toString(), "ok4")
        }
