// vybe-test: kotlin/kotlin_string_starts_with_ops/test_ends_with_case
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_starts_with_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "Hello"
            __check((s.endsWith("O", true).toString()).toString(), "true")
        }
