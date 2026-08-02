// vybe-test: kotlin/type_cast_edges/test_is_on_null_value
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any? = null
            __check((value is String).toString(), "false")
        }
