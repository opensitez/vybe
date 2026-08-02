// vybe-test: kotlin/loops/test_nested_for_break_affects_inner_loop_only
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var total = 0
            for (i in 1..3) {
                for (j in 1..5) {
                    if (j == 4) break
                    total += i * j
                }
            }
            println(total)
        }

