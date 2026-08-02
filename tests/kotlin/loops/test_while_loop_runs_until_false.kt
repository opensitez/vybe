// vybe-test: kotlin/loops/test_while_loop_runs_until_false
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var i = 1
            var out = 0
            while (i <= 4) {
                out += i
                i += 1
            }
            println(out)
        }

