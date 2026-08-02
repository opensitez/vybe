// vybe-test: kotlin/for_loop_variants/test_for_outer_inner_product
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var out = 0
            for (i in 1..3) {
                for (j in 1..4) {
                    out += if ((i + j) % 2 == 0) 1 else 0
                }
            }
            println(out)
        }

