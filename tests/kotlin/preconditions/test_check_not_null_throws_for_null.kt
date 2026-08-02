// vybe-test: kotlin/preconditions/test_check_not_null_throws_for_null
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun main() {
            try {
                checkNotNull<String>(null)
                println("no")
            } catch (e: IllegalStateException) {
                println("thrown")
            }
        }

