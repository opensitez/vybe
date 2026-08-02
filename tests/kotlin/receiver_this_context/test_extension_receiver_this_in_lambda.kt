// vybe-test: kotlin/receiver_this_context/test_extension_receiver_this_in_lambda
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

fun String.wrap(): String = this.also { __check(("start").toString(), "start") } + "!"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("ok".wrap()).toString(), "ok!")
        }
