// vybe-test: kotlin/collections_maps_ops/test_map_iterator_is_fail_fast_when_mutated
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            val iter = map.iterator()
            println(iter.hasNext())
            println(iter.next().key)
            map["c"] = 3
            try {
                iter.next()
                println("no_fail")
            } catch (e: ConcurrentModificationException) {
                println("fail_fast")
            }
            println(map.size)
        }

