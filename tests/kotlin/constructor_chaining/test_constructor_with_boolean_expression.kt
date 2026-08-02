// vybe-test: kotlin/constructor_chaining/test_constructor_with_boolean_expression
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Tri(val v: Int) {
            val ok: Boolean
            init { ok = v % 2 == 0 }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Tri(4).ok).toString(), "true")
            __check((Tri(5).ok).toString(), "false")
        }
