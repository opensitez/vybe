// vybe-test: kotlin/collection_projection_views/test_map_entry_set_mutation
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val entries = map.entries
            val it = entries.iterator()
            while (it.hasNext()) {
                val e = it.next()
                if (e.key == "a") {
                    it.remove()
                }
            }
            println(map.size)
            println(map.containsKey("a"))
            println(map["b"])
        }

