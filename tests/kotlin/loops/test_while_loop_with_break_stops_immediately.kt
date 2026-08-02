// vybe-test: kotlin/loops/test_while_loop_with_break_stops_immediately
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var i = 0
            var out = 0
            while (true) {
                if (i == 3) break
                out += i
                i += 1
            }
            println(out)
        }

