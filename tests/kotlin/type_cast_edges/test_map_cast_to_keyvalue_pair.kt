// vybe-test: kotlin/type_cast_edges/test_map_cast_to_keyvalue_pair
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = mapOf("a" to 1)
            val cast = value as? Map<String, Int>
            __check((cast?.size ?: -1).toString(), "1")
        }
