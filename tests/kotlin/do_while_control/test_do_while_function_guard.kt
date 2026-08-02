// vybe-test: kotlin/do_while_control/test_do_while_function_guard
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun keep(i: Int): Boolean = i < 3
        fun main() {
            var i = 0
            var out = ""
            do {
                out += i.toString()
                i += 1
            } while (keep(i))
            println(out)
        }

