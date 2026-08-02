// vybe-test: kotlin/constructor_chaining/test_constructor_secondary_calls_base_super
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

open class Parent(val a: Int)
        class Child : Parent {
            constructor() : super(1)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Child().a).toString(), "1")
        }
