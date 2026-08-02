// vybe-test: kotlin/type_cast_edges/test_list_generic_cast_with_star_projection
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = listOf("a", "b")
            val list = value as? List<*>
            __check((list?.size ?: -1).toString(), "2")
        }
