// vybe-test: kotlin/for_loop_variants/test_for_down_to_range
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var out = ""
            for (i in 5 downTo 2) {
                out += i.toString()
            }
            println(out)
        }

