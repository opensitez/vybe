// vybe-test: kotlin/constructor_chaining/test_constructor_property_with_expr_init
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Calc(val a: Int) {
            val b = a + 1
            val c = b + a
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Calc(4)
            __check((c.b).toString(), "5")
            __check((c.c).toString(), "9")
        }
