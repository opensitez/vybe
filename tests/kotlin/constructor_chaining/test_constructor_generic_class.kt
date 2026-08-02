// vybe-test: kotlin/constructor_chaining/test_constructor_generic_class
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Holder<T>(val v: T) {
            val text = v.toString()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder("x")
            __check((h.text).toString(), "x")
        }
