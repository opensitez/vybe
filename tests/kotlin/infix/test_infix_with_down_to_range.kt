// vybe-test: kotlin/infix/test_infix_with_down_to_range
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun main() {
            var sum = 0
            for (i in 5 downTo 1) {
                sum += i
            }
            println(sum)
        }

