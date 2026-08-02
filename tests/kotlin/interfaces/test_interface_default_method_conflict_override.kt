// vybe-test: kotlin/interfaces/test_interface_default_method_conflict_override
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface A {
            fun label(): String = "A"
        }
        interface B {
            fun label(): String = "B"
        }
        class C : A, B {
            override fun label(): String = super<A>.label() + "+" + super<B>.label()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: A = C()
            __check((value.label()).toString(), "A+B")
        }
