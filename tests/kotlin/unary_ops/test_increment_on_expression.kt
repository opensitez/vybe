// vybe-test: kotlin/unary_ops/test_increment_on_expression
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 1
            val y = ++x + x++
            __check((y).toString(), "3")
            __check((x).toString(), "3")
        }
