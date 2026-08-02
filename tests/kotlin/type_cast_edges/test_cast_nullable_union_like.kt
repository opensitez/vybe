// vybe-test: kotlin/type_cast_edges/test_cast_nullable_union_like
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Any? = null
            val b: String? = a as? String
            __check((b == null).toString(), "true")
        }
