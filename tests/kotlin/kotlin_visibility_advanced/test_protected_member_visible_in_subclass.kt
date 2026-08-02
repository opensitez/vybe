// vybe-test: kotlin/kotlin_visibility_advanced/test_protected_member_visible_in_subclass
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

open class Base {
            protected fun token() = "ok"
        }

        class Child : Base() {
            fun reveal() = token()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Child().reveal()).toString(), "ok")
        }
