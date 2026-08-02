// vybe-test: kotlin/inheritance_dispatch/test_abstract_overrides_are_required_by_concrete_class
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

abstract class Base {
            abstract val title: String
        }

        class Leaf : Base() {
            override val title: String = "leaf"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Leaf().title).toString(), "leaf")
        }
