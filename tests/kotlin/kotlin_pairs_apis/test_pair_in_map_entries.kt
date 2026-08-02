// vybe-test: kotlin/kotlin_pairs_apis/test_pair_in_map_entries
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun main() {
            val values = mapOf(1 to "one", 2 to "two")
            var out = ""
            for (entry in values) {
                out += entry.key.toString() + entry.value
            }
            println(out)
        }

