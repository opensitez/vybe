// vybe-test: kotlin/lambdas/test_lambda_as_map_operation
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun transform(v: Int, op: (Int) -> Int): Int { return op(v) }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((transform(3, { it + 2 })).toString(), "5")
__check((transform(4, { x -> x * 2 })).toString(), "8") }
