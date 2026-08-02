// vybe-test: kotlin/loops/test_for_over_empty_collection_keeps_accumulator_zero
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            val values = intArrayOf()
            var seen = 0
            for (value in values) {
                seen += value
            }
            println(seen)
        }

