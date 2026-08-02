// vybe-test: kotlin/extension_properties/test_extension_property_boolean_to_word
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Boolean.word: String get() = if (this) "yes" else "no"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true.word).toString(), "yes")
            __check((false.word).toString(), "no")
        }
