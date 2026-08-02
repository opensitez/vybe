// vybe-test: kotlin/for_loop_variants/test_for_range_membership_inside_loop
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var out = 0
            for (i in 1..10) {
                if (i in 4..6) out += i
            }
            println(out)
        }

