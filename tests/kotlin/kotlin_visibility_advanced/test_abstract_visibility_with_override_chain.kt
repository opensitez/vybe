// vybe-test: kotlin/kotlin_visibility_advanced/test_abstract_visibility_with_override_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

abstract class Root {
            protected abstract fun token(): String
        }

        class Leaf : Root() {
            override fun token() = "seen"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Leaf().token()).toString(), "seen")
        }
