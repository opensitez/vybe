// vybe-test: kotlin/do_while_control/test_do_while_in_function
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun sum(n: Int): Int {
            var i = 0
            var out = 0
            do {
                out += i
                i++
            } while (i < n)
            return out
        }
        fun main() {
            println(sum(4))
        }

