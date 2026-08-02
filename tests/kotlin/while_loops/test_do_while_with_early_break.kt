// vybe-test: kotlin/while_loops/test_do_while_with_early_break
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            var out = 0
            do {
                if (i == 3) break
                out += i
                i += 1
            } while (true)
            println(out)
        }

