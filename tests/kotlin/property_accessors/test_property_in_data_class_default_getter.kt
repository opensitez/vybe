// vybe-test: kotlin/property_accessors/test_property_in_data_class_default_getter
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

data class Point(val x: Int, val y: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Point(1, 2)
            __check((p.x).toString(), "1")
            __check((p.y).toString(), "2")
        }
