// vybe-test: kotlin/extension_functions/test_extension_property_with_nullable_receiver
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

val Int?.orZero: Int
            get() = this ?: 0

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left: Int? = null
            val right: Int? = 12
            __check((left.orZero).toString(), "0")
            __check((right.orZero).toString(), "12")
        }
