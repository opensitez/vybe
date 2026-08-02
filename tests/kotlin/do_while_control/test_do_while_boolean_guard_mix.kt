// vybe-test: kotlin/do_while_control/test_do_while_boolean_guard_mix
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun shouldRun(v: Int): Boolean = v % 2 == 0
        fun main() {
            var i = 0
            var out = ""
            do {
                if (shouldRun(i)) out += i.toString()
                i += 1
            } while (i < 5)
            println(out)
        }

