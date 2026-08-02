// vybe-test: kotlin/for_loop_variants/test_for_shadowing_boundaries
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var out = 0
            for (i in 1..3) {
                val i = i + 10
                out += i
            }
            println(out)
        }

