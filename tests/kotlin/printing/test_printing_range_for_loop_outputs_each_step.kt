// vybe-test: kotlin/printing/test_printing_range_for_loop_outputs_each_step
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun main() {
            var output = ""
            for (i in 1..3) {
                output += i.toString() + ","
            }
            println(output)
        }

