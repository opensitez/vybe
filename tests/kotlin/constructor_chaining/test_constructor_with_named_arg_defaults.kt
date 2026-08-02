// vybe-test: kotlin/constructor_chaining/test_constructor_with_named_arg_defaults
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Namer(val first: String = "x", val last: String = "y")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Namer(last = "z")
            __check((a.first).toString(), "x")
            __check((a.last).toString(), "z")
        }
