// vybe-test: kotlin/unary_ops/test_postfix_increment
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 1
            __check((x++).toString(), "1")
            __check((x).toString(), "2")
        }
