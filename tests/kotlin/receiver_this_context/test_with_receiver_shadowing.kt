// vybe-test: kotlin/receiver_this_context/test_with_receiver_shadowing
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

data class Holder(val value: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = Holder("inner").run {
                val value = "local"
                __check((value).toString(), "local")
                this.value
            }
            __check((out).toString(), "inner")
        }
