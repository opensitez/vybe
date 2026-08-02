// vybe-test: kotlin/preconditions/test_check_in_variant_logic
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun parseAge(age: Int?): Int {
            checkNotNull(age)
            check(age >= 0)
            return age
        }

        fun main() {
            try {
                println(parseAge(-1))
            } catch (e: IllegalStateException) {
                println("state")
            }
        }

