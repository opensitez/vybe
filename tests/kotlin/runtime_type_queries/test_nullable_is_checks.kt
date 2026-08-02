// vybe-test: kotlin/runtime_type_queries/test_nullable_is_checks
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v: String? = null
            __check((v is String).toString(), "false")
            __check((v == null).toString(), "true")
            val w: Any? = null
            __check((w is String?).toString(), "true")
            __check((w is Int?).toString(), "true")
        }
