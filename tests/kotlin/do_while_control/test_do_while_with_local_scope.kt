// vybe-test: kotlin/do_while_control/test_do_while_with_local_scope
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var total = 0
            do {
                val step = 2
                total += step
            } while (total < 8)
            println(total)
        }

