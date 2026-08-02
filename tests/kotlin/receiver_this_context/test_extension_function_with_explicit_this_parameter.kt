// vybe-test: kotlin/receiver_this_context/test_extension_function_with_explicit_this_parameter
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

class Holder {
            fun Int.addToHolder(): Int = this + 10
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            __check((h.run {
                3.addToHolder()
            }).toString(), "13")
        }
