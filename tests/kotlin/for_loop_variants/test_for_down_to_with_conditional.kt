// vybe-test: kotlin/for_loop_variants/test_for_down_to_with_conditional
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var out = 0
            for (i in 10 downTo 1) {
                if (i % 4 == 0) out += i
            }
            println(out)
        }

