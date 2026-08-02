// vybe-test: kotlin/properties/test_property_setter_records_previous_state
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Logarithm {
            private var raw = 0
            var value: Int
                get() = raw
                set(next) { raw = next + 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Logarithm()
            value.value = 3
            value.value = 7
            __check((value.value).toString(), "8")
        }
