// vybe-test: kotlin/kotlin_string_starts_with_ops/test_ends_with_offset
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_starts_with_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abcXYZ"
            val sub = "XYZ"
            __check((s.endsWith(sub).toString()).toString(), "true")
        }
