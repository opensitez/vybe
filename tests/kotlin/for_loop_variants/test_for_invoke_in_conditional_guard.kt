// vybe-test: kotlin/for_loop_variants/test_for_invoke_in_conditional_guard
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun isOdd(x: Int): Boolean = x % 2 == 1
        fun main() {
            var out = 0
            for (i in 1..10) {
                if (isOdd(i)) out += i
            }
            println(out)
        }

