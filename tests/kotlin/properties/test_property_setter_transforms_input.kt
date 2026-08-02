// vybe-test: kotlin/properties/test_property_setter_transforms_input
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Clamp {
            private var raw: Int = 0
            var value: Int
                get() = raw
                set(next) { raw = next * 10 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Clamp()
            c.value = 3
            __check((c.value).toString(), "30")
        }
