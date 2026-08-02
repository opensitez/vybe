// vybe-test: kotlin/collections_maps/test_map_value_view_tracks_updates
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            val values = map.values
            map["a"] = 4
            map["c"] = 3
            var sum = 0
            for (value in values) {
                sum += value
            }
            println(sum)
            map.remove("b")
            println(values.size)
            println(map.size)
        }

