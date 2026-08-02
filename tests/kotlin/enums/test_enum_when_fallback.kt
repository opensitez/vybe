// vybe-test: kotlin/enums/test_enum_when_fallback
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Mode { ON, OFF }

        fun describe(mode: Mode): String {
            return when (mode) {
                Mode.ON -> "enabled"
                Mode.OFF -> "disabled"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(Mode.ON)).toString(), "enabled")
            __check((describe(Mode.OFF)).toString(), "disabled")
        }
