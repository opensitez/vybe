// vybe-test: kotlin/property_accessors/test_property_custom_interface_implementation
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

interface HasValue { var value: Int }
        class Box(override var value: Int) : HasValue
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box(8)
            __check((b.value).toString(), "8")
            b.value = 9
            __check((b.value).toString(), "9")
        }
