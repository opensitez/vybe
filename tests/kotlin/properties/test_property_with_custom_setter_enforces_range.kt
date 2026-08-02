// vybe-test: kotlin/properties/test_property_with_custom_setter_enforces_range
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class RangeTracker {
            var value: Int = 0
                set(next) {
                    field = if (next < 0) 0 else next
                }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val tracker = RangeTracker()
            tracker.value = -1
            __check((tracker.value).toString(), "0")
            tracker.value = 3
            __check((tracker.value).toString(), "3")
        }
