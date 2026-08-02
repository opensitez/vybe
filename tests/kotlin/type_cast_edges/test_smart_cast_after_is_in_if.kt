// vybe-test: kotlin/type_cast_edges/test_smart_cast_after_is_in_if
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "hello"
            val out = if (value is String) {
                value.length
            } else {
                -1
            }
            __check((out).toString(), "5")
        }
