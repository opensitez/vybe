// vybe-test: kotlin/non_local_returns/test_local_return_from_lambda_with_label
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun sumEven(values: List<Int>): Int {
            var total = 0
            values.forEach {
                if (it % 2 == 0) return@forEach
                total += it
            }
            return total
        }

        fun main() {
            println(sumEven(listOf(1, 2, 3, 4, 5)))
        }

