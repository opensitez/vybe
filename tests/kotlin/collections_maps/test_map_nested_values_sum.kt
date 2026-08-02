// vybe-test: kotlin/collections_maps/test_map_nested_values_sum
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val buckets = mapOf(
                "left" to listOf(1, 2, 3),
                "right" to listOf(4, 5)
            )
            var total = 0
            for (value in buckets["left"]!!) {
                total += value
            }
            for (value in buckets["right"]!!) {
                total += value
            }
            println(total)
        }

