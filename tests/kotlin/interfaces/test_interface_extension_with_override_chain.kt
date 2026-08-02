// vybe-test: kotlin/interfaces/test_interface_extension_with_override_chain
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Parent {
            fun base(): String = "base"
        }

        interface Child : Parent {
            override fun base(): String = "child"
        }

        class Impl : Child

        fun Parent.tag(): String = this.base() + ":tagged"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c: Child = Impl()
            __check((c.tag()).toString(), "child:tagged")
            __check(((c as Parent).base()).toString(), "child")
            __check((c.base()).toString(), "child")
        }
