// vybe-test: kotlin/kotlin_string_starts_with_ops/test_ends_with_basic
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_starts_with_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "language"
            __check((s.endsWith("age").toString()).toString(), "true")
            __check((s.endsWith("lang").toString()).toString(), "false")
        }
