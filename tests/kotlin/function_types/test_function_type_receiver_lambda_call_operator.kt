// vybe-test: kotlin/function_types/test_function_type_receiver_lambda_call_operator
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun applyBlock(v: Int, op: Int.() -> Int): Int {
            return v.op()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val inc = fun Int.() -> Int { return this + 1 }
            __check((applyBlock(5, inc)).toString(), "6")
        }
