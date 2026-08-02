// vybe-test: kotlin/loops/test_for_loop_skips_empty_down_to_range
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var seen = 0
            for (i in 1 downTo 3) {
                seen += i
            }
            println(seen)
        }

