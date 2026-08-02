// vybe-test: kotlin/strings/test_string_repeat_negative_throws
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun main() {
            try {
                println("x".repeat(-1))
            } catch (e: Exception) {
                println("repeat-error")
            }
        }

