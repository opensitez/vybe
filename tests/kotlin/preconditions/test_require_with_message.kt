// vybe-test: kotlin/preconditions/test_require_with_message
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun main() {
            try {
                require(false, { "bad" })
                println("no")
            } catch (e: IllegalArgumentException) {
                println(e.message)
            }
        }

