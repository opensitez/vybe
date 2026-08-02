// vybe-test: kotlin/collections_maps/test_map_iteration_keys_values
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val data = mapOf("x" to 1, "y" to 2)
            var keys = ""
            var sum = 0
            for (entry in data.entries) {
                keys += entry.key
                sum += entry.value
            }
            println(keys)
            println(sum)
        }

