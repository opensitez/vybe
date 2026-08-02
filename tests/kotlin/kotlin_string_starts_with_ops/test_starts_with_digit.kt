// vybe-test: kotlin/kotlin_string_starts_with_ops/test_starts_with_digit
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_starts_with_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "123abc"
            __check((s.startsWith("1").toString()).toString(), "true")
            __check((s.endsWith("abc").toString()).toString(), "true")
        }
