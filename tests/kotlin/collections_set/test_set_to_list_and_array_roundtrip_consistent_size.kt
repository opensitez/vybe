// vybe-test: kotlin/collections_set/test_set_to_list_and_array_roundtrip_consistent_size
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3)
            val list = values.toList()
            val array = values.toTypedArray()
            __check((list.size).toString(), "3")
            __check((array.size).toString(), "3")
            __check((list.contains(2)).toString(), "true")
            __check((array.size == list.size).toString(), "true")
        }
