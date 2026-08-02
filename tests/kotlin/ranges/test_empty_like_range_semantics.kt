// vybe-test: kotlin/ranges/test_empty_like_range_semantics
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var seen = 0
            for (i in 3 until 3) {
                seen += i
            }
            println(seen)
        }

