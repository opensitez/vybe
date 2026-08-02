// vybe-test: kotlin/collections_maps/test_map_entries_aggregation
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val metrics = mapOf("read" to 5, "write" to 7, "update" to 3)
            var total = 0
            var hasUpdate = false
            for ((name, value) in metrics) {
                total += value
                if (name == "update") {
                    hasUpdate = true
                }
            }
            println(total)
            println(hasUpdate)
            println(metrics.size)
        }

