// vybe-test: kotlin/inheritance_dispatch/test_multiple_interface_implementations_can_override_both
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface A {
            fun tag(): String = "A"
        }

        interface B {
            fun tag(): String = "B"
        }

        class C : A, B {
            override fun tag(): String = super<A>.tag() + "+" + super<B>.tag()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((C().tag()).toString(), "A+B")
        }
