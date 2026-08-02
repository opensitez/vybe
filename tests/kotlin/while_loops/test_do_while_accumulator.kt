// vybe-test: kotlin/while_loops/test_do_while_accumulator
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            var out = 0
            do {
                out += i
                i += 1
            } while (i < 4)
            println(out)
            println(i)
        }

