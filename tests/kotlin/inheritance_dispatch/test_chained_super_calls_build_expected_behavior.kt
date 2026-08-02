// vybe-test: kotlin/inheritance_dispatch/test_chained_super_calls_build_expected_behavior
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open fun label(): String = "base"
        }

        open class Mid : Base() {
            override fun label(): String = super.label() + ":mid"
        }

        class Leaf : Mid() {
            override fun label(): String = super.label() + ":leaf"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Base = Leaf()
            __check((value.label()).toString(), "base:mid:leaf")
        }
