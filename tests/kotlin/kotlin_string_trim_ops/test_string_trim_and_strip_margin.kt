// vybe-test: kotlin/kotlin_string_trim_ops/test_string_trim_and_strip_margin
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_trim_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("  ab ".trim()).toString(), "ab")
            __check(("|a\n|b".trimMargin("|")).toString(), "a\nb")
        }
