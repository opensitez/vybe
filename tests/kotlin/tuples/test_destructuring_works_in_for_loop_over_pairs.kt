// vybe-test: kotlin/tuples/test_destructuring_works_in_for_loop_over_pairs
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun main() {
            val rows = listOf(Pair("a", 1), Pair("b", 2), Pair("c", 3))
            var total = ""
            for ((label, value) in rows) {
                total += "$label$value-"
            }
            println(total)
        }

