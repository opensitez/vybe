// vybe-test: kotlin/collections_maps/test_map_get_value_throws_when_missing_key
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val scores = mapOf("a" to 1, "b" to 2)
            try {
                println(scores.getValue("z"))
            } catch (e: NoSuchElementException) {
                println("missing")
            }
        }

