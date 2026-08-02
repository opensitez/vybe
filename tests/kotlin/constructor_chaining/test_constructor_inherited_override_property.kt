// vybe-test: kotlin/constructor_chaining/test_constructor_inherited_override_property
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

open class Base { open val v = 1 }
        class Sub : Base() { override val v = 2 }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b: Base = Sub()
            __check((b.v).toString(), "2")
        }
