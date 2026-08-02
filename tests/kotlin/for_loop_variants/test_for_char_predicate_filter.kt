// vybe-test: kotlin/for_loop_variants/test_for_char_predicate_filter
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var out = ""
            for (ch in 'a'..'f') {
                if (ch != 'd') out += ch
            }
            println(out)
        }

