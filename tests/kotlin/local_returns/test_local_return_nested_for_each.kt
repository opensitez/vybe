// vybe-test: kotlin/local_returns/test_local_return_nested_for_each
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            val result = StringBuilder()
            outer@ for (r in 1..2) {
                inner@ for (c in 1..3) {
                    if (r == 1 && c == 2) continue@outer
                    result.append(r).append(c)
                }
            }
            println(result.toString())
        }

