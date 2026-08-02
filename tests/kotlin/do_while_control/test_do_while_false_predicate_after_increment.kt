// vybe-test: kotlin/do_while_control/test_do_while_false_predicate_after_increment
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = 0
            do {
                out += i
                i += 2
            } while (i > 10)
            println(out)
        }

