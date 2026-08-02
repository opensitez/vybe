// vybe-test: kotlin/control_flow/test_char_range_for_loop
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var letters = ""
            for (c in 'a'..'d') {
                letters += c
            }
            println(letters)
        }

