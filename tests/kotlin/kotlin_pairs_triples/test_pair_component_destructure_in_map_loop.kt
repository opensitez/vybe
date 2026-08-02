// vybe-test: kotlin/kotlin_pairs_triples/test_pair_component_destructure_in_map_loop
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun main() {
            val map = mapOf("x" to 10, "y" to 20)
            var total = 0
            for ((k, v) in map) {
                total += if (k == "x") v else 0
            }
            println(total)
        }

