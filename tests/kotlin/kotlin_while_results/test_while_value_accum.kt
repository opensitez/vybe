// vybe-test: kotlin/kotlin_while_results/test_while_value_accum
// origin: languages/kotlin/tests/kotlin/test_kotlin_while_results.rs

fun main() {
            var i = 0
            var out = 0
            while (i < 3) {
                out = out + i
                i = i + 1
            }
            println(out)
        }

