// vybe-test: kotlin/unary_ops/test_unary_boolean_algebra
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = true
            val b = false
            val out = ! (a || b) == false
            __check((out).toString(), "true")
        }
