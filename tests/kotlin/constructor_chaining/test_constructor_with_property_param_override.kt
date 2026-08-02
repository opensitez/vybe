// vybe-test: kotlin/constructor_chaining/test_constructor_with_property_param_override
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

open class Base(val v: Int)
        class Derived(v: Int) : Base(v + 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Derived(4).v).toString(), "5")
        }
