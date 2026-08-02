// vybe-test: kotlin/kotlin_result_runtime/test_result_on_success_and_on_failure_have_side_effects
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var successSeen = false
            var failureSeen = false
            runCatching { 5 }
                .onSuccess { successSeen = true }
                .onFailure { failureSeen = true }
            runCatching<Int> { throw RuntimeException("fail") }
                .onSuccess { successSeen = true }
                .onFailure { failureSeen = true }
            __check((successSeen).toString(), "true")
            __check((failureSeen).toString(), "true")
        }
