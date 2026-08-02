// vybe-test: kotlin/kotlin_visibility_advanced/test_public_open_override_is_visible_via_reference
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

open class Parent { open val name: String = "p" }
        class Child : Parent() { override val name: String = "c" }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base: Parent = Child()
            __check((base.name).toString(), "c")
        }
