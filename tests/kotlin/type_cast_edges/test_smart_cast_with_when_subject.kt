// vybe-test: kotlin/type_cast_edges/test_smart_cast_with_when_subject
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = listOf(1, 2)
            val out = when (value) {
                is List<*> -> value.size
                else -> -1
            }
            __check((out).toString(), "2")
        }
