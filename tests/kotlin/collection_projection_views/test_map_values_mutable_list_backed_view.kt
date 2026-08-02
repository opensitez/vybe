// vybe-test: kotlin/collection_projection_views/test_map_values_mutable_list_backed_view
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("x" to 1, "y" to 2)
            val values = map.values
            __check((values.sum()).toString(), "3")
            map["x"] = 9
            __check((values.joinToString(",")).toString(), "9,2")
            val copied = values.toMutableList()
            copied.remove(2)
            __check((values.size).toString(), "2")
            __check((copied.size).toString(), "1")
        }
