// vybe-test: kotlin/type_cast_edges/test_cast_to_nullable_type
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = null
            val any: Any? = value
            val value2 = any as? String
            __check((value2 == null).toString(), "true")
        }
