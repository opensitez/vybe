// vybe-test: kotlin/comparison_ops/test_compare_with_char_codes
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('a'.code < 'b'.code).toString(), "true")
            __check(('z'.code > 'y'.code).toString(), "true")
        }
