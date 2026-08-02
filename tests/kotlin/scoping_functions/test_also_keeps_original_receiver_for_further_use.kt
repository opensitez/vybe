// vybe-test: kotlin/scoping_functions/test_also_keeps_original_receiver_for_further_use
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Holder(var text: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Holder("x")
                .also { it.text = it.text + "y" }
                .also { it.text = it.text + "z" }
            __check((item.text).toString(), "xyz")
        }
