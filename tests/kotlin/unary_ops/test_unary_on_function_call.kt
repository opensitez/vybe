// vybe-test: kotlin/unary_ops/test_unary_on_function_call
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun value(): Int = 4
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((-value()).toString(), "-4")
            __check((+value()).toString(), "4")
        }
