// vybe-test: kotlin/kotlin_do_while_results/test_do_while_minimum_runs
// origin: languages/kotlin/tests/kotlin/test_kotlin_do_while_results.rs

fun main() {
            var i = 5
            var out = 0
            do {
                out = out + i
                i = i - 1
            } while (i > 5)
            println(out)
        }

