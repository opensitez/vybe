// vybe-test: kotlin/collection_projection_views/test_map_key_set_view_reflects_updates
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val keys = map.keys
            val values = map.values
            __check((keys.joinToString(",")).toString(), "a,b")
            __check((values.joinToString(",")).toString(), "1,2")
            map["c"] = 3
            map["a"] = 9
            __check((keys.joinToString(",")).toString(), "a,b,c")
            __check((values.joinToString(",")).toString(), "9,2,3")
        }
