// vybe-test: kotlin/loops/test_for_on_char_range_with_step_includes_expected_codepoints
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var text = ""
            for (c in 'a'..'f' step 2) {
                text += c.toString()
            }
            println(text)
        }

