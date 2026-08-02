// vybe-test: kotlin/inheritance_dispatch/test_super_calls_parent_implementation
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = super.label() + ":child"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Child().label()).toString(), "base:child")
        }
