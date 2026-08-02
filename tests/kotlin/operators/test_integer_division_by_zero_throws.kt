// vybe-test: kotlin/operators/test_integer_division_by_zero_throws
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun main() {
            try {
                println(10 / 0)
            } catch (e: Exception) {
                println("caught")
            }
        }

