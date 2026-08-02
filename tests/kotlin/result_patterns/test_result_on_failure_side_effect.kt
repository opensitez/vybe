// vybe-test: kotlin/result_patterns/test_result_on_failure_side_effect
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun main() {
            val value = Result.failure<String>(Exception("bad")).onSuccess { println("hit") }
            value.onFailure { println("fail") }
            println(value.isFailure)
        }

