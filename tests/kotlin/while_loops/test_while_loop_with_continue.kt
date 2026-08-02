// vybe-test: kotlin/while_loops/test_while_loop_with_continue
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            var out = 0
            while (i < 6) {
                i += 1
                if (i % 2 == 0) continue
                out += i
            }
            println(out)
        }

