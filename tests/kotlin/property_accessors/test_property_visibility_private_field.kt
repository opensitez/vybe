// vybe-test: kotlin/property_accessors/test_property_visibility_private_field
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            private var _v = 1
            var v: Int
                get() = _v
                set(value) { _v = value }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            b.v = 12
            __check((b.v).toString(), "12")
        }
