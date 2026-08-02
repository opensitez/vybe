// vybe-test: kotlin/preconditions/test_check_with_message
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun main() {
            try {
                check(false, { "state bad" })
                println("no")
            } catch (e: IllegalStateException) {
                println(e.message)
            }
        }

