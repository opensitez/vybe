// vybe-test: kotlin/kotlin_destructuring_maps/test_mutable_map_entry_rewrite_with_destructure
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun main() {
            val map = mutableMapOf("a" to 1)
            for ((k, _) in map) {
                map[k] = 5
            }
            println(map["a"])
        }

