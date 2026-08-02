// vybe-test: kotlin/for_loop_variants/test_for_nested_outer_uses
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var out = ""
            for (row in 1..2) {
                for (col in 1..2) {
                    out += "${row}${col} "
                }
            }
            println(out.trim())
        }

