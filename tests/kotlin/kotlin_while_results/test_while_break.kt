// vybe-test: kotlin/kotlin_while_results/test_while_break
// origin: languages/kotlin/tests/kotlin/test_kotlin_while_results.rs

fun main() {
            var i = 0
            while (true) {
                if (i == 2) {
                    break
                }
                i = i + 1
            }
            println(i)
        }

