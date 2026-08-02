// vybe-test: kotlin/local_returns/test_local_return_continue_loop
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            val out = StringBuilder()
            for (i in 1..4) {
                if (i == 2) continue
                out.append(i)
            }
            println(out.toString())
        }

