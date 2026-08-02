// vybe-test: kotlin/extension_properties/test_extension_property_array_first_or_empty
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Array<Int>.firstOrNullSafe: Int get() = this.firstOrNull() ?: -1
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((intArrayOf().firstOrNullSafe).toString(), "-1")
            __check((intArrayOf(1, 4).firstOrNullSafe).toString(), "1")
        }
