// vybe-test: kotlin/local_returns/test_local_return_lambda_named
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            fun run(items: List<Int>): String {
                val r = StringBuilder()
                items.forEach loop@{
                    if (it == 2) return@loop
                    r.append(it)
                }
                return r.toString()
            }
            println(run(listOf(1, 2, 3)))
        }

