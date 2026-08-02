// vybe-test: kotlin/runtime_type_queries/test_generic_type_checks_star_projection
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: Any = listOf(1, 2, 3)
            __check((values is List<*>).toString(), "true")
            __check((values is Set<*>).toString(), "false")
        }
