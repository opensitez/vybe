// vybe-test: kotlin/collections_maps/test_set_to_list_roundtrip_keeps_distinct_elements
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mutableListOf(1, 2, 2, 3, 1)
            val unique = source.toSet()
            val back = unique.toMutableList()
            back.add(4)
            __check((unique.size).toString(), "3")
            __check((back.size).toString(), "4")
            __check((back.contains(4)).toString(), "true")
            __check((source.size).toString(), "5")
        }
