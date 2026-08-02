// vybe-test: kotlin/type_cast_edges/test_array_as_list_safecast
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = intArrayOf(1, 2, 3)
            val cast = value as? List<Int>
            __check((cast == null).toString(), "true")
        }
