// vybe-test: kotlin/collections_maps/test_mutable_list_insert_and_remove_at_index
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3)
            values.add(1, 9)
            values.removeAt(2)
            __check((values.size).toString(), "3")
            __check((values[1]).toString(), "9")
            __check((values[2]).toString(), "3")
        }
