// vybe-test: kotlin/extension_functions/test_extension_function_dispatch_uses_static_receiver_type
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

open class Base
        class Child : Base()

        fun Base.label(): String = "base"
        fun Child.label(): String = "child"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val static_base: Base = Child()
            __check((static_base.label()).toString(), "base")
            __check((Child().label()).toString(), "child")
        }
