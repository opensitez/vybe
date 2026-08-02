// vybe-test: kotlin/for_loop_variants/test_for_range_var_bounds
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var start = 2
            var end = 5
            var total = 0
            for (i in start..end) {
                total += i
            }
            println(total)
        }

