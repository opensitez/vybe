// vybe-test: kotlin/kotlin_string_starts_with_ops/test_starts_with_basic
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_starts_with_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "kotlin"
            __check((s.startsWith("kot").toString()).toString(), "true")
            __check((s.startsWith("lin").toString()).toString(), "false")
        }
