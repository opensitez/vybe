// vybe-test: kotlin/inheritance_dispatch/test_subclass_without_override_uses_base_behavior
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open fun text(): String = "base"
        }

        class Direct : Base()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Direct().text()).toString(), "base")
        }
