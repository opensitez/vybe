// vybe-test: kotlin/properties/test_computed_property_uses_other_state
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Holder {
            var left = 2
            var right = 5
            val total: Int
                get() = left + right
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Holder()
            __check((value.total).toString(), "7")
            value.left = 7
            __check((value.total).toString(), "12")
        }
