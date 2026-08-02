// vybe-test: kotlin/for_loop_variants/test_for_large_step
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var total = 0
            for (i in 0..20 step 5) {
                total += i
            }
            println(total)
        }

