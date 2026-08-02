// vybe-test: kotlin/extension_properties/test_extension_property_list_size_text
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val List<Int>.sizeText: String get() = when (size) {
            0 -> "empty"
            1 -> "single"
            else -> "many"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf<Int>().sizeText).toString(), "empty")
            __check((listOf(1).sizeText).toString(), "single")
            __check((listOf(1, 2).sizeText).toString(), "many")
        }
