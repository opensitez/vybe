// vybe-test: kotlin/extension_functions/test_extension_for_companion_object_acts_like_factory
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

class Factory private constructor(val value: Int) {
            companion object
        }

        fun Factory.Companion.from(value: Int): Factory = Factory(value)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Factory.from(7).value).toString(), "7")
        }
