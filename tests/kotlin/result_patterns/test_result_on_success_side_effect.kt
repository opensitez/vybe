// vybe-test: kotlin/result_patterns/test_result_on_success_side_effect
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun main() {
            val value = Result.success("ok").onSuccess { println("hit") }
            value.onFailure { println("no") }
            println(value.getOrNull())
        }

