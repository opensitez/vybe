// vybe-test: kotlin/kotlin_constructor_parameters/test_constructor_parameter_property
// origin: languages/kotlin/tests/kotlin/test_kotlin_constructor_parameters.rs

class Box(val payload: String) {
            fun valueLabel(): String {
                return "<" + payload + ">"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box("x")
            __check((b.valueLabel()).toString(), "<x>")
        }
