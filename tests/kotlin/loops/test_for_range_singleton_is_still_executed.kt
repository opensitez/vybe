// vybe-test: kotlin/loops/test_for_range_singleton_is_still_executed
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var count = 0
            for (i in 7 downTo 7) {
                count += i
            }
            println(count)
        }

