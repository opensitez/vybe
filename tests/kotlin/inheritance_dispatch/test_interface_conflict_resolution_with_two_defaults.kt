// vybe-test: kotlin/inheritance_dispatch/test_interface_conflict_resolution_with_two_defaults
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface A {
            fun text(): String = "a"
        }

        interface B {
            fun text(): String = "b"
        }

        class C : A, B {
            override fun text(): String = super<A>.text() + "," + super<B>.text()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((C().text()).toString(), "a,b")
        }
