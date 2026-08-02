// vybe-test: kotlin/kotlin_do_while_results/test_do_while_with_continue
// origin: languages/kotlin/tests/kotlin/test_kotlin_do_while_results.rs

fun main() {
            var i = 0
            var out = ""
            do {
                i = i + 1
                if (i == 2) {
                    continue
                }
                out = out + i.toString()
            } while (i < 4)
            println(out)
        }

