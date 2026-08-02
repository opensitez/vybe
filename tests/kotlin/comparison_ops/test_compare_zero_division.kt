// vybe-test: kotlin/comparison_ops/test_compare_zero_division
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun main() {
            try {
                println(1 == (1 / 0))
            } catch (e: Exception) {
                println("err")
            }
        }

