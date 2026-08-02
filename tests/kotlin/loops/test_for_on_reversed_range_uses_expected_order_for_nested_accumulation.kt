// vybe-test: kotlin/loops/test_for_on_reversed_range_uses_expected_order_for_nested_accumulation
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var sequence = ""
            for (i in 3 downTo 1) {
                for (j in 0..1) {
                    sequence += i
                    sequence += ":"
                }
            }
            println(sequence)
        }

