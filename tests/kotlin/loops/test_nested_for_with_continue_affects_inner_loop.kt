// vybe-test: kotlin/loops/test_nested_for_with_continue_affects_inner_loop
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var total = 0
            for (i in 1..3) {
                for (j in 1..5) {
                    if (j == 3) continue
                    total += i + j
                }
            }
            println(total)
        }

