// vybe-test: kotlin/type_cast_edges/test_array_cast_preserves_elements_type
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = arrayOf(1, 2, 3)
            val cast = value as? Array<Int>
            __check((cast?.size ?: -1).toString(), "3")
        }
