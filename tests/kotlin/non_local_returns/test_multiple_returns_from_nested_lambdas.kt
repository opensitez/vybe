// vybe-test: kotlin/non_local_returns/test_multiple_returns_from_nested_lambdas
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun firstDivisible(values: List<Int>): Int {
            values.forEach {
                if (it % 3 == 0) {
                    return it
                }
            }
            return -1
        }

        fun all(values: List<Int>): Int {
            values.forEach { first ->
                if (first == 0) return 0
                if (first > 0) return@forEach
            }
            return 9
        }

        fun main() {
            println(firstDivisible(listOf(2, 5, 6, 7)))
            println(all(listOf(-1, 1, 2)))
        }

