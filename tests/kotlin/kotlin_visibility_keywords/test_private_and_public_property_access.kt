// vybe-test: kotlin/kotlin_visibility_keywords/test_private_and_public_property_access
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_keywords.rs

open class Base {
            private val secret = "secret"
            public val shown = "shown"
            protected open val inherited = "inherited"
        }

        class Child : Base() {
            override val inherited: String = "childInherited"
            fun exposeInherited(): String {
                return inherited
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Child()
            __check((b.shown).toString(), "shown")
            __check((b.exposeInherited()).toString(), "childInherited")
        }
