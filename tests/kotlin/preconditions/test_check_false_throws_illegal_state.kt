// vybe-test: kotlin/preconditions/test_check_false_throws_illegal_state
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun main() {
            try {
                check(false)
                println("no")
            } catch (e: IllegalStateException) {
                println(e::class.simpleName)
            }
        }

