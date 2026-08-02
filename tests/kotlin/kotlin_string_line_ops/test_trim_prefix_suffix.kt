// vybe-test: kotlin/kotlin_string_line_ops/test_trim_prefix_suffix
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("   padded   ".trim()).toString(), "padded")
            __check(("...x".trimStart('.')).toString(), "x")
            __check(("...x".trimEnd('.')).toString(), "...x")
            __check(("   x".trimStart()).toString(), "x")
        }
