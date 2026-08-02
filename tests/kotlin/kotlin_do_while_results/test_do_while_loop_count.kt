// vybe-test: kotlin/kotlin_do_while_results/test_do_while_loop_count
// origin: languages/kotlin/tests/kotlin/test_kotlin_do_while_results.rs

fun main() {
            var i = 0
            var out = ""
            do {
                out = out + i.toString()
                i = i + 1
            } while (i < 3)
            println(out)
        }

