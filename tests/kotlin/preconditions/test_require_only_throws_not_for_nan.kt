// vybe-test: kotlin/preconditions/test_require_only_throws_not_for_nan
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun main() {
            try {
                require(Double.NaN.isFinite())
                println("finite")
            } catch (e: IllegalArgumentException) {
                println("bad")
            }
        }

