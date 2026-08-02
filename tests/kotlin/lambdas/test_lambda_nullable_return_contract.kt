// vybe-test: kotlin/lambdas/test_lambda_nullable_return_contract
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun resolve(flag: Boolean): (Int) -> Int? {
    return if (flag) {
        { x -> x + 1 }
    } else {
        { _ -> null }
    }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val f = resolve(true)
    val g = resolve(false)
    __check((f(1) ?: -1).toString(), "2")
    __check((g(1) ?: -1).toString(), "-1")
}
