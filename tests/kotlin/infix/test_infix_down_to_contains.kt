// vybe-test: kotlin/infix/test_infix_down_to_contains
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun main() {
            var total = 0
            for (n in 4 downTo 1) {
                if (n in 3 downTo 1) {
                    total += n
                }
            }
            println(total)
        }

