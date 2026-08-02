// vybe-test: kotlin/constructor_chaining/test_constructor_captures_param_in_init
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Capture(val x: Int) {
            val y: Int
            init { y = x * 3 }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Capture(4).y).toString(), "12")
        }
