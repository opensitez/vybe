// vybe-test: kotlin/local_returns/test_local_return_in_while_conditional
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun main() {
            var i = 0
            while (i < 4) {
                i += 1
                if (i == 2) continue
                if (i == 4) break
                println(i)
            }
        }

