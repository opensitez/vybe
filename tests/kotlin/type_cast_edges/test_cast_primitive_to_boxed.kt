// vybe-test: kotlin/type_cast_edges/test_cast_primitive_to_boxed
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Int = 3
            val boxed: Any = value
            val cast = boxed as Int
            __check((cast + 1).toString(), "4")
        }
