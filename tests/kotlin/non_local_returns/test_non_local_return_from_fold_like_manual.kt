// vybe-test: kotlin/non_local_returns/test_non_local_return_from_fold_like_manual
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun firstLarge(values: List<Int>): Int {
            values.forEach {
                if (it > 10) {
                    return it
                }
            }
            return -10
        }

        fun main() {
            println(firstLarge(listOf(3, 11, 12)))
            println(firstLarge(listOf(1, 2, 3)))
        }

