// vybe-test: kotlin/inheritance_dispatch/test_virtual_method_is_dispatched_from_parent_reference
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = "child"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Base = Child()
            __check((item.label()).toString(), "child")
        }
