// vybe-test: kotlin/collections_maps/test_mutable_map_update_does_not_reorder_existing_keys
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val state = linkedMapOf("first" to 1, "second" to 2)
            state["first"] = 9
            state["first"] = 11
            var keys = ""
            for ((key, _) in state) {
                keys += key
            }
            println(keys)
            println(state["first"])
            println(state.size)
        }

