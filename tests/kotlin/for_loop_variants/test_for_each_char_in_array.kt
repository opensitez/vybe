// vybe-test: kotlin/for_loop_variants/test_for_each_char_in_array
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            val values = charArrayOf('a', 'b', 'c')
            var out = ""
            for (c in values) {
                out += c
            }
            println(out)
        }

