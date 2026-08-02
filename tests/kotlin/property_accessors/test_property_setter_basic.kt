// vybe-test: kotlin/property_accessors/test_property_setter_basic
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var v: Int = 0
                set(value) { field = value }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            b.v = 4
            __check((b.v).toString(), "4")
        }
