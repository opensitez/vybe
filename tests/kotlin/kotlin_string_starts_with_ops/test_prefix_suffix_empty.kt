// vybe-test: kotlin/kotlin_string_starts_with_ops/test_prefix_suffix_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_starts_with_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = ""
            __check((s.startsWith("x").toString()).toString(), "false")
            __check((s.endsWith("").toString()).toString(), "true")
        }
