// vybe-test: kotlin/lambdas/test_lambda_with_receiver_style_extension_call
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val shout: String.() -> String = { this.uppercase() + "!" }
    __check(("go".shout()).toString(), "GO!")
}
