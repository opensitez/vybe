// vybe-test: kotlin/collections_maps/test_map_of_linked_preserves_assignment_insertion_order
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val history = linkedMapOf("first" to 1, "second" to 2, "third" to 3)
            var keys = ""
            for ((key, _) in history) {
                keys += key
            }
            println(keys)
            history["second"] = 4
            var keysAfter = ""
            for ((key, _) in history) {
                keysAfter += key
            }
            println(keysAfter)
        }

