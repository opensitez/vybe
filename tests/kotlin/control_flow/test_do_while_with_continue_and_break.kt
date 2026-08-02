// vybe-test: kotlin/control_flow/test_do_while_with_continue_and_break
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var i = 0
            var out = ""
            do {
                i++
                if (i == 2) continue
                if (i > 4) break
                out += i.toString()
            } while (i < 6)
            println(out)
        }

