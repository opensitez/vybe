// vybe-test: kotlin/for_loop_variants/test_for_over_while_style_pattern
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var i = 1
            var out = 0
            for (x in 1..10) {
                if (i > 5) break
                out += x
                i += 1
            }
            println(out)
        }

