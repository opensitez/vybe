// vybe-test: kotlin/lambdas/test_lambda_nested_call
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun make(): (Int) -> Int {
            return { value -> value + 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val fnRef = make()
            __check((fnRef(6)).toString(), "7")
            __check((make()(3)).toString(), "4")
        }
