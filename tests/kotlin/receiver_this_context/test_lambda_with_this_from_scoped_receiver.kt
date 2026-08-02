// vybe-test: kotlin/receiver_this_context/test_lambda_with_this_from_scoped_receiver
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

data class Box(val value: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = Box("x").run {
                with(this) {
                    __check((value).toString(), "x")
                    this.value.length
                }
            }
            __check((out).toString(), "1")
        }
