// vybe-test: kotlin/class_delegation/test_class_delegation_multiple_interface_members_forwarding
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface A { fun a(): String }
        interface B { fun b(): String }

        class Impl : A, B {
            override fun a() = "A"
            override fun b() = "B"
        }

        class Wrapper(private val impl: Impl) : A by impl, B by impl

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val w = Wrapper(Impl())
            __check((w.a()).toString(), "A")
            __check((w.b()).toString(), "B")
        }
