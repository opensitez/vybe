// vybe-test: kotlin/constructor_chaining/test_constructor_with_default_lambda
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Maker(val f: () -> Int = { 3 }) {
            fun value() = f()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Maker().value()).toString(), "3")
            __check((Maker { 5 } .value()).toString(), "5")
        }
