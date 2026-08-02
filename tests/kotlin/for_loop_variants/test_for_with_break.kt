// vybe-test: kotlin/for_loop_variants/test_for_with_break
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var total = 0
            for (i in 1..10) {
                if (i == 4) break
                total += i
            }
            println(total)
        }

