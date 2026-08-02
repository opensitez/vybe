// vybe-test: kotlin/do_while_control/test_do_while_with_mutable_flag
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = ""
            var running = true
            do {
                if (i >= 3) running = false
                out += i.toString()
                i += 1
            } while (running)
            println(out)
        }

