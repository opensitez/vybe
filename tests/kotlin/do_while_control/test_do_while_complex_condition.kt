// vybe-test: kotlin/do_while_control/test_do_while_complex_condition
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun shouldRun(x: Int): Boolean = x <= 2
        fun main() {
            var i = 0
            var total = 0
            do {
                total += i
                i += 1
            } while (i < 6 && shouldRun(i))
            println(total)
            println(i)
        }

