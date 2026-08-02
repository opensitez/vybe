// vybe-test: kotlin/extension_functions/test_extension_property_can_be_computed_multiple_times
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

var Int.squareCount: Int
            get() = this * this
            set(value) {}

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 4
            __check((value.squareCount).toString(), "16")
            __check((value.squareCount).toString(), "16")
        }
