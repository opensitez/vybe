// vybe-test: kotlin/loops/test_while_loop_with_boolean_flag_and_loop_variable
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var active = true
            var i = 0
            var total = 0
            while (active) {
                total += i
                i += 1
                active = i < 5
            }
            println(total)
        }

