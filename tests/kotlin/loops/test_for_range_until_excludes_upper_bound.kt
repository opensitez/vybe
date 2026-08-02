// vybe-test: kotlin/loops/test_for_range_until_excludes_upper_bound
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var total = 0
            for (i in 1 until 4) {
                total += i
            }
            println(total)
        }

