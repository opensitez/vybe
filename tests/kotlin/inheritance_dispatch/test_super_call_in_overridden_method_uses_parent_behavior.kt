// vybe-test: kotlin/inheritance_dispatch/test_super_call_in_overridden_method_uses_parent_behavior
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open fun score(value: Int): Int = value * 2
        }

        class Child : Base() {
            override fun score(value: Int): Int = super.score(value) + 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Child().score(3)).toString(), "7")
        }
