// vybe-test: kotlin/comparison_ops/test_compare_char
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('a' < 'c').toString(), "true")
            __check(('z' > 'm').toString(), "true")
            __check(('a' == 'a').toString(), "true")
        }
