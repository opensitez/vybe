// vybe-test: kotlin/kotlin_property_accessors_advanced/test_property_reassigns_and_getter_validation
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class RangeValue {
            private var _value = 0
            var value: Int
                get() = _value
                set(v) { _value = if (v > 10) 10 else v }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = RangeValue()
            r.value = 15
            __check((r.value).toString(), "10")
        }
