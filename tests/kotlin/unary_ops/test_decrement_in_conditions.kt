// vybe-test: kotlin/unary_ops/test_decrement_in_conditions
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 4
            val out = if (--x > 2) "gt" else "lte"
            __check((out).toString(), "gt")
        }
