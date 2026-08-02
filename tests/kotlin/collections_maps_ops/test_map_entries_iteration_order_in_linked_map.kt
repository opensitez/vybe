// vybe-test: kotlin/collections_maps_ops/test_map_entries_iteration_order_in_linked_map
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun main() {
            val map = linkedMapOf("first" to 1, "second" to 2, "third" to 3)
            var keys = ""
            var sum = 0
            for (entry in map.entries) {
                keys += entry.key
                sum += entry.value
            }
            println(keys)
            println(sum)
        }

