// vybe-test: kotlin/type_cast_edges/test_as_safe_cast_wrong_type_returns_null
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any? = 10
            val text = value as? String
            __check((text == null).toString(), "true")
        }
