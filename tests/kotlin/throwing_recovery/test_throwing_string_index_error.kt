// vybe-test: kotlin/throwing_recovery/test_throwing_string_index_error
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun main() {
            try {
                val text = "abc"
                println(text[99])
            } catch (e: Exception) {
                println("error")
            }
        }

