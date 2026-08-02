// vybe-test: kotlin/type_cast_edges/test_safe_cast_of_null_to_string
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any? = null
            val text = value as? String
            __check((text).toString(), "null")
        }
