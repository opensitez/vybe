// vybe-test: kotlin/kotlin_string_starts_with_ops/test_starts_with_offset
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_starts_with_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "prefix"
            __check((s.startsWith("fix", startIndex = 3).toString()).toString(), "true")
        }
