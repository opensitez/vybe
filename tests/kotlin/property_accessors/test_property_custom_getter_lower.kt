// vybe-test: kotlin/property_accessors/test_property_custom_getter_lower
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Point(val x: Int, val y: Int) {
            val min: Int get() = if (x < y) x else y
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Point(2, 7).min).toString(), "2")
        }
