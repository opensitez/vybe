// vybe-test: kotlin/collections_maps_ops/test_map_for_each_accumulation
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            var sumKeys = 0
            var sumValues = 0
            map.forEach { _, value ->
                sumValues += value
            }
            for (_ in map) {
                sumKeys += 1
            }
            println(sumValues)
            println(sumKeys)
        }

