// vybe-test: kotlin/lambdas/test_higher_order_function
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun execute(a: Int, b: Int, op: (Int, Int) -> Int): Int {
            return op(a, b)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sum = execute(15, 25, { x, y -> x + y })
            __check((sum).toString(), "40")
        }
