// vybe-test: kotlin/extension_properties/test_extension_property_nested_list_count
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val List<List<Int>>.flattenedCount: Int get() = this.sumOf { it.size }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf(listOf(1, 2), listOf(3)).flattenedCount).toString(), "3")
        }
