// vybe-test: kotlin/property_accessors/test_property_setter_normalized
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var value: Int = 0
                set(value) { field = if (value < 0) 0 else value }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            b.value = -2
            __check((b.value).toString(), "0")
            b.value = 7
            __check((b.value).toString(), "7")
        }
