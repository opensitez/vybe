// vybe-test: kotlin/runtime_type_queries/test_is_check_on_primitives_vs_wrappers
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Any = 5
            val b: Any = "kotlin"
            __check((a is Int).toString(), "true")
            __check((a is String).toString(), "false")
            __check((b is String).toString(), "true")
            __check((b is Number).toString(), "false")
        }
