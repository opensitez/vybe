// vybe-test: kotlin/collections_maps_ops/test_map_get_value_throws_when_missing
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun main() {
            val map = mapOf("a" to 1)
            try {
                println(map.getValue("b"))
            } catch (e: NoSuchElementException) {
                println("missing")
            }
        }

