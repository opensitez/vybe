// vybe-test: kotlin/lambdas/test_lambda_with_destructured_pair_param
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun transform(input: Pair<Int, Int>, op: (Int, Int) -> Int): Int {
    return op(input.first, input.second)
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    __check((transform(Pair(3, 4), { a, b -> a + b })).toString(), "7")
    __check((transform(Pair(10, 2), { a, b -> a - b })).toString(), "8")
}
