// vybe-test: kotlin/collections_maps/test_map_entries_iterator_mutation_is_fail_fast_for_structure_changes
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val source = mutableMapOf("a" to 1, "b" to 2)
            val iter = source.entries.iterator()
            println(iter.hasNext())
            println(iter.next().key)
            source["c"] = 3
            try {
                iter.next()
                println("no_fail")
            } catch (e: ConcurrentModificationException) {
                println("fail_fast")
            }
        }

