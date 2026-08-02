// vybe-test: kotlin/inheritance_dispatch/test_multiple_class_levels_share_virtual_method
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open fun route(): String = "base"
        }

        open class Mid : Base() {
            override fun route(): String = "mid"
        }

        class Leaf : Mid() {
            override fun route(): String = "leaf"
        }

        fun emit(route: Base): String = route.route()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((emit(Base())).toString(), "base")
            __check((emit(Mid())).toString(), "mid")
            __check((emit(Leaf())).toString(), "leaf")
        }
