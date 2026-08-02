// vybe-test: kotlin/printing/test_printing_joined_output_across_multiple_calls
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun main() {
            val a = listOf("a", "b", "c")
            for (value in a) {
                print(value)
                print("-")
            }
            println("done")
        }

