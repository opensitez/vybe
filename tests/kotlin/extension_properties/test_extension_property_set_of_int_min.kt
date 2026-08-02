// vybe-test: kotlin/extension_properties/test_extension_property_set_of_int_min
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Set<Int>.minOrMinusOne: Int get() = this.minOrNull() ?: -1
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((setOf<Int>().minOrMinusOne).toString(), "-1")
            __check((setOf(9, 2, 5).minOrMinusOne).toString(), "2")
        }
