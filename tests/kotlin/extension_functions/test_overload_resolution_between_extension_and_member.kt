// vybe-test: kotlin/extension_functions/test_overload_resolution_between_extension_and_member
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

class Box {
            fun value(): Int = 1
        }

        fun Box.value(): Int = 4

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().value()).toString(), "1")
        }
