// vybe-test: kotlin/preconditions/test_require_false_throws_illegal_argument
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun main() {
            try {
                require(false)
                println("no")
            } catch (e: IllegalArgumentException) {
                println(e::class.simpleName)
            }
        }

