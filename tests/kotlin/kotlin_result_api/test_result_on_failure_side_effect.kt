// vybe-test: kotlin/kotlin_result_api/test_result_on_failure_side_effect
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var marker = ""
            runCatching<Int> { throw IllegalArgumentException("x") }
                .onSuccess { marker += "ok" }
                .onFailure { marker += it::class.simpleName.toString() }
            __check((marker).toString(), "IllegalArgumentException")
        }
