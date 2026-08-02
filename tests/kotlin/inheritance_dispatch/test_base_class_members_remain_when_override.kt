// vybe-test: kotlin/inheritance_dispatch/test_base_class_members_remain_when_override
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = "child"
            fun asBase(): Base = this
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val child = Child()
            __check((child.label()).toString(), "child")
            __check((child.asBase().label()).toString(), "child")
        }
