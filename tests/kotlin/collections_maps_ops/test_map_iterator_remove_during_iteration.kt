// vybe-test: kotlin/collections_maps_ops/test_map_iterator_remove_during_iteration
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 3, "d" to 4)
            val iter = map.entries.iterator()
            while (iter.hasNext()) {
                val current = iter.next()
                if (current.value % 2 == 0) {
                    iter.remove()
                }
            }
            println(map.size)
            println(map["b"])
            println(map["d"])
            println(map["a"] + map["c"])
        }

