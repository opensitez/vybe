// vybe-test: kotlin/inheritance_dispatch/test_override_chain_across_multiple_levels
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open fun route(): String = "base"
        }

        open class Mid : Base() {
            override fun route(): String = "mid" + super.route()
        }

        class Leaf : Mid() {
            override fun route(): String = "leaf" + super.route()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val node: Base = Leaf()
            __check((node.route()).toString(), "leafmidbase")
        }
