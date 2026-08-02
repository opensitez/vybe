// vybe-test: kotlin/type_cast_edges/test_as_safe_cast_matches_type
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any? = "hello"
            val text = value as? String
            __check((text).toString(), "hello")
        }
