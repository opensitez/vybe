// vybe-test: kotlin/for_loop_variants/test_for_empty_range
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var total = 0
            for (i in 5 until 5) {
                total += i
            }
            println(total)
        }

