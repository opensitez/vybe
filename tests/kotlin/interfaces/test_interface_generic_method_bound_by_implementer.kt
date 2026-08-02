// vybe-test: kotlin/interfaces/test_interface_generic_method_bound_by_implementer
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Formatter {
            fun <T : Number> format(value: T): String
        }

        class IntFormatter : Formatter {
            override fun <T : Number> format(value: T): String = "n:" + value.toInt().toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f: Formatter = IntFormatter()
            __check((f.format(12)).toString(), "n:12")
            __check((f.format(12.4)).toString(), "n:12")
        }
