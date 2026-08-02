// vybe-test: kotlin/kotlin_visibility_advanced/test_visibility_of_extension_target
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

private class Core {
            fun value() = 9
        }

        fun Core.expose() = value()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Core().expose()).toString(), "9")
        }
