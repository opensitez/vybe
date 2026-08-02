// vybe-test: kotlin/kotlin_range_repetition/test_int_range_step_expressions
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_repetition.rs

fun main() {
            var acc = ""
            for (i in 0..6 step 2) {
                acc += i.toString()
            }
            println(acc)
        }

