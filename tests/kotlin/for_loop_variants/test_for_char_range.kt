// vybe-test: kotlin/for_loop_variants/test_for_char_range
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var out = ""
            for (ch in 'a'..'e') {
                out += ch.toString()
            }
            println(out)
        }

