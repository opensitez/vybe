// vybe-test: kotlin/do_while_control/test_do_while_mixed_types_in_condition
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var sum = 0L
            do {
                sum += i.toLong()
                i += 1
            } while (i.toLong() < 4L)
            println(sum)
        }

