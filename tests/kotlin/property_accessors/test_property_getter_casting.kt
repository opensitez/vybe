// vybe-test: kotlin/property_accessors/test_property_getter_casting
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Holder {
            val number: Number = 2
            val intValue: Int get() = number.toInt()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder().intValue).toString(), "2")
        }
