// vybe-test: kotlin/properties/test_property_nullable_with_setter_default_handling
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Holder {
            private var raw: String? = null
            var value: String?
                get() = raw
                set(next) { raw = next ?: "" }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            h.value = null
            __check(("[" + h.value + "]").toString(), "[]")
        }
