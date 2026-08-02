// vybe-test: kotlin/collections_maps/test_list_index_lookup_and_last_position
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val values = listOf(5, 6, 7, 6, 8)
            println(values.indexOf(6))
            println(values.lastIndexOf(6))
            var output = ""
            for (i in values.indices) {
                if (i % 2 == 1) {
                    output += values[i].toString()
                }
            }
            println(values.size)
            println(output)
        }

