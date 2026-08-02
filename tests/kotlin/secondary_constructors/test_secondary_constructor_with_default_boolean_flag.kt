// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_with_default_boolean_flag
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Marker {
            val active: Boolean
            val label: String

            constructor(label: String) {
                this.label = label
                this.active = false
            }

            constructor(label: String, active: Boolean) : this(label) {
                if (active) this.label = label + "!"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Marker("x")
            val b = Marker("x", true)
            __check((a.active).toString(), "false")
            __check((a.label).toString(), "x")
            __check((b.label).toString(), "x!")
        }
