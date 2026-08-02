// vybe-test: kotlin/loops/test_for_loop_over_char_range
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var text = ""
            for (c in 'a'..'d') {
                text += c
            }
            println(text)
        }

