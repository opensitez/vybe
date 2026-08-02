// vybe-test: kotlin/while_loops/test_while_with_string_accumulator
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            var text = ""
            while (i < 3) {
                text += i.toString()
                i += 1
            }
            println(text)
        }

