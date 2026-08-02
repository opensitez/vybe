// vybe-test: kotlin/kotlin_string_trim_ops/test_string_starts_with_and_ends_with
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_trim_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "prefix:value"
            __check((s.startsWith("pre")).toString(), "true")
            __check((s.endsWith("value")).toString(), "true")
        }
