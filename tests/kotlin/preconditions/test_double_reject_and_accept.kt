// vybe-test: kotlin/preconditions/test_double_reject_and_accept
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun main() {
            for (value in listOf(-1, 1)) {
                val ok = try {
                    require(value > 0)
                    "yes"
                } catch (e: IllegalArgumentException) {
                    "no"
                }
                println(ok)
            }
        }

