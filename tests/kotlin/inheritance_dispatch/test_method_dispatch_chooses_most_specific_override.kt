// vybe-test: kotlin/inheritance_dispatch/test_method_dispatch_chooses_most_specific_override
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
            val value: Base = Child()
            __check((value.label()).toString(), "child")
        }
