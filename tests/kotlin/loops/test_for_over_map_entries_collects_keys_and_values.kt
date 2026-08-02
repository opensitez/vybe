// vybe-test: kotlin/loops/test_for_over_map_entries_collects_keys_and_values
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            val values = mapOf("a" to 1, "b" to 2)
            var keys = ""
            var sum = 0
            for (entry in values.entries) {
                keys += entry.key
                sum += entry.value
            }
            println(keys)
            println(sum)
        }

