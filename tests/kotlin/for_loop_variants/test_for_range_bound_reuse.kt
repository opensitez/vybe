// vybe-test: kotlin/for_loop_variants/test_for_range_bound_reuse
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var start = 2
            var end = 6
            var out = 0
            for (i in start until end) {
                out += i
            }
            println(out)
        }

