// vybe-test: kotlin/do_while_control/test_do_while_nested
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var outer = 0
            var inner = 0
            do {
                outer += 1
                var j = 0
                do {
                    inner += j
                    j += 1
                } while (j < 2)
            } while (outer < 3)
            println(outer)
            println(inner)
        }

