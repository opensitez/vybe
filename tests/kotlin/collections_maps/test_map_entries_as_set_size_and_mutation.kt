// vybe-test: kotlin/collections_maps/test_map_entries_as_set_size_and_mutation
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val inventory = mutableMapOf("x" to 1, "y" to 2)
            val entryView = inventory.entries
            var hasXOne = false
            for (entry in entryView) {
                if (entry.key == "x" && entry.value == 1) {
                    hasXOne = true
                }
            }
            println(entryView.size)
            inventory["z"] = 3
            println(entryView.size)
            println(hasXOne)
        }

