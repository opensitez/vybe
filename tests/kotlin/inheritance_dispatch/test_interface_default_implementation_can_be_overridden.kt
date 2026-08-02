// vybe-test: kotlin/inheritance_dispatch/test_interface_default_implementation_can_be_overridden
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface Messenger {
            fun text(): String = "default"
        }

        class Custom : Messenger {
            override fun text(): String = "custom"
        }

        class InheritDefault : Messenger

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Custom().text()).toString(), "custom")
            __check((InheritDefault().text()).toString(), "default")
        }
