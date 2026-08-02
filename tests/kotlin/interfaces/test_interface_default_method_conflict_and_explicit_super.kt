// vybe-test: kotlin/interfaces/test_interface_default_method_conflict_and_explicit_super
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Left {
            fun label(): String = "left"
        }

        interface Right {
            fun label(): String = "right"
        }

        class Both : Left, Right {
            override fun label(): String = super<Left>.label()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val both: Left = Both()
            __check((both.label()).toString(), "left")
            __check(((both as Right).label()).toString(), "left")
        }
