// vybe-test: kotlin/infix/test_infix_with_step_and_accumulation
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun main() {
            var sum = 0
            for (x in 0..10 step 3) {
                sum += x
            }
            println(sum)
        }

