// vybe-test: kotlin/constructor_chaining/test_constructor_inheritance_chain
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

open class A(val a: Int)
        class B(x: Int) : A(x + 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((B(2).a).toString(), "3")
        }
