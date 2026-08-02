// vybe-test: kotlin/property_accessors/test_property_backing_field_private_setter
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            private var _v = 0
            var value: Int
                get() = _v
                private set
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            __check((b.value).toString(), "0")
        }
