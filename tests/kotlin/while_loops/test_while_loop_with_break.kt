// vybe-test: kotlin/while_loops/test_while_loop_with_break
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            var out = 0
            while (i < 10) {
                i += 1
                if (i == 4) break
                out += i
            }
            println(i)
            println(out)
        }

