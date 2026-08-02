// vybe-test: kotlin/do_while_control/test_do_while_string_counter
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = ""
            do {
                out += "#" + i.toString()
                i += 1
            } while (i < 3)
            println(out)
        }

