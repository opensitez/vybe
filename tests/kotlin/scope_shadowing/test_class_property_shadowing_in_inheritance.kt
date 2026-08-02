// vybe-test: kotlin/scope_shadowing/test_class_property_shadowing_in_inheritance
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

open class Parent {
            val value = "parent"
        }
        class Child : Parent() {
            val value = "child"
            fun reveal() = super.value + ":" + value
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Child()
            __check((c.value).toString(), "child")
            __check((c.reveal()).toString(), "parent:child")
        }
