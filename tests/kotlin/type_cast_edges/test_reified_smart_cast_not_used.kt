// vybe-test: kotlin/type_cast_edges/test_reified_smart_cast_not_used
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

inline fun <reified T> isType(value: Any): Boolean = value is T

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isType<String>("x")).toString(), "true")
            __check((isType<Int>("x")).toString(), "false")
        }
