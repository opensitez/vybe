// vybe-test: kotlin/do_while_control/test_do_while_on_function_result
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun tick(x: Int): Boolean = x < 2
        fun main() {
            var i = 0
            var out = 0
            do {
                out += i
                i++
            } while (tick(i))
            println(out)
            println(i)
        }

