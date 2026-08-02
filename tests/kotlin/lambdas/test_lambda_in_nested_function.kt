// vybe-test: kotlin/lambdas/test_lambda_in_nested_function
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun factory(base: Int): (Int) -> Int { return { v -> v + base } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val addTen = factory(10)
__check((addTen(5)).toString(), "15") }
