// vybe-test: kotlin/ranges/test_range_step_with_start_end_expressions
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun build(start: Int, end: Int, step: Int): String {
            var out = ""
            for (value in start..end step step) {
                out += value.toString()
            }
            return out
        }

        fun main() {
            println(build(1, 7, 2))
        }

