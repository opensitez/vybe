// vybe-test: kotlin/type_cast_edges/test_is_check_true_and_false
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x: Any = 7
            __check((x is Int).toString(), "true")
            __check((x is String).toString(), "false")
        }
