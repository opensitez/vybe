// vybe-test: kotlin/property_accessors/test_property_with_custom_equals
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box(val value: Int) {
            val isPositive: Boolean get() = value > 0
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Box(3)
            val b = Box(-1)
            __check((a.isPositive).toString(), "true")
            __check((b.isPositive).toString(), "false")
        }
