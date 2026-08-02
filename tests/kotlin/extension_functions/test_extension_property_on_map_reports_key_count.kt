// vybe-test: kotlin/extension_functions/test_extension_property_on_map_reports_key_count
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

val Map<String, Int>.keyText: String
            get() = keys.joinToString("|")

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            __check((values.keyText).toString(), "a|b|c")
            __check((values.keyText).toString(), "a|b|c")
        }
