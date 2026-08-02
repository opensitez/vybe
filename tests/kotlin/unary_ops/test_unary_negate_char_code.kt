// vybe-test: kotlin/unary_ops/test_unary_negate_char_code
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = 'a'.code
            __check((c).toString(), "97")
            __check((-c).toString(), "-97")
        }
