// vybe-test: kotlin/companion_objects/test_companion_method_uses_its_own_properties
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Calculator {
            companion object {
                private const val scale = 10
                fun scaled(value: Int): Int = value * scale
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Calculator.scaled(3)).toString(), "30")
        }
