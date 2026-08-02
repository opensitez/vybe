// vybe-test: kotlin/for_loop_variants/test_for_with_continue
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var total = 0
            for (i in 1..6) {
                if (i % 2 == 0) continue
                total += i
            }
            println(total)
        }

