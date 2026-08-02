// vybe-test: kotlin/boolean_logic/test_boolean_with_while_break_continue_controls
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun main() {
            var i = 0
            var total = 0
            while (i < 6) {
                i++
                if (i % 2 == 0) {
                    continue
                }
                if (i > 4) {
                    break
                }
                total += i
            }
            println(i)
            println(total)
        }

