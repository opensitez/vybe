// vybe-test: kotlin/collections_maps/test_mutable_list_insert_and_update
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val values = mutableListOf("a", "c")
            values.add(1, "b")
            values[2] = "d"
            var output = ""
            for (value in values) {
                output += value
            }
            println(output)
            println(values.size)
        }

