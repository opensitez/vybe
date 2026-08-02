// vybe-test: kotlin/constructor_chaining/test_constructor_nested_class_call
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Outer {
            class Inner(val v: Int)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val i = Outer.Inner(9)
            __check((i.v).toString(), "9")
        }
