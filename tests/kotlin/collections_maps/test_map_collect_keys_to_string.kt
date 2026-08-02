// vybe-test: kotlin/collections_maps/test_map_collect_keys_to_string
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val data = mapOf("a" to 10, "b" to 20, "c" to 30)
            var keys = ""
            for (k in data.keys) {
                keys += k
            }
            println(keys)
            println(data.keys.size)
        }

