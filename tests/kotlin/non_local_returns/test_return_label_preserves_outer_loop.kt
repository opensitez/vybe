// vybe-test: kotlin/non_local_returns/test_return_label_preserves_outer_loop
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun main() {
            var count = 0
            outer@ for (i in 1..4) {
                listOf(1, 2).forEach {
                    if (i == 3 && it == 1) return@outer
                    count += i
                }
            }
            println(count)
        }

