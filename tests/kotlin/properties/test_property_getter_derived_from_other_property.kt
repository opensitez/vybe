// vybe-test: kotlin/properties/test_property_getter_derived_from_other_property
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Square(val side: Int) {
            val area: Int
                get() = side * side
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Square(5).area).toString(), "25")
        }
