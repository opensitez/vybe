// vybe-test: kotlin/local_returns/test_local_return_named_label_in_run
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            val result = run label@{
                listOf(1, 2).forEach {
                    if (it == 1) return@label "first"
                }
                "none"
            }
            println(result)
        }

