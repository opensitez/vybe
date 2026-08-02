// vybe-test: kotlin/result_patterns/test_result_get_or_throw_failure_throws
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun main() {
            val value = runCatching { throw IllegalArgumentException("bad") }
            try {
                value.getOrThrow()
                println("ok")
            } catch (e: IllegalArgumentException) {
                println(e.message)
            }
        }

