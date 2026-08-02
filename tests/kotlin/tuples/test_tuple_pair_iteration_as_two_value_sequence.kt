// vybe-test: kotlin/tuples/test_tuple_pair_iteration_as_two_value_sequence
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun main() {
            val data = listOf(Pair(1, 10), Pair(2, 20), Pair(3, 30))
            var total = 0
            for (i in data) {
                total += i.first
                total += i.second
            }
            println(total)
        }

