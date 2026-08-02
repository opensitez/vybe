// vybe-test: kotlin/kotlin_map_apis/test_map_entries_iteration
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun main() {
            val map = linkedMapOf("x" to 1, "y" to 2)
            var seen = ""
            for (entry in map.entries) {
                seen += entry.key + entry.value
            }
            println(seen)
        }

