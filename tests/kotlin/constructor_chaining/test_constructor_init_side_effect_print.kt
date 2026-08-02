// vybe-test: kotlin/constructor_chaining/test_constructor_init_side_effect_print
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Side {
            init { __check(("ok").toString(), "ok") }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Side()
        }
