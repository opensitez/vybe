// vybe-test: kotlin/kotlin_result_api/test_result_on_success_then_failure_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var count = 0
            runCatching { 7 }.onSuccess { count += it }
            runCatching<Int> { throw RuntimeException("bad") }.onFailure { count += 5 }
            __check((count).toString(), "12")
        }
