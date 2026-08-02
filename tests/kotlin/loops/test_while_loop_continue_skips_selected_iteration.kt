// vybe-test: kotlin/loops/test_while_loop_continue_skips_selected_iteration
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var i = 0
            var total = 0
            while (i < 6) {
                i += 1
                if (i % 2 == 0) continue
                total += i
            }
            println(total)
        }

