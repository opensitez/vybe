// vybe-test: kotlin/control_flow/test_while_with_postcondition_and_early_exit
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var i = 1
            var total = 0
            while (i <= 10) {
                total += i
                if (total >= 6) break
                i += 1
            }
            println(total)
            println(i)
        }

