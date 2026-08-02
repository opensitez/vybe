// vybe-test: kotlin/local_returns/test_local_return_breakout_loop
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            val out = StringBuilder()
            for (i in 1..5) {
                if (i == 4) break
                out.append(i)
            }
            println(out.toString())
        }

