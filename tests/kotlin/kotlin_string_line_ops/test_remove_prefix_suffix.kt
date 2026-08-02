// vybe-test: kotlin/kotlin_string_line_ops/test_remove_prefix_suffix
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("prefix:value".removePrefix("prefix:")).toString(), "value")
            __check(("prefix:value".removePrefix("x")).toString(), "prefix:value")
            __check(("value/suffix".removeSuffix("/suffix")).toString(), "value")
            __check(("value/suffix".removeSuffix("x")).toString(), "value/suffix")
        }
