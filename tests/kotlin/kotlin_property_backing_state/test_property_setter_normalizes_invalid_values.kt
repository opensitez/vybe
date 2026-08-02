// vybe-test: kotlin/kotlin_property_backing_state/test_property_setter_normalizes_invalid_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_backing_state.rs

class Metric {
            private var _count: Int = 0
            var count: Int
                get() = _count
                set(value) {
                    _count = if (value < 0) 0 else value
                }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = Metric()
            m.count = -5
            __check((m.count).toString(), "0")
            m.count = 8
            __check((m.count).toString(), "8")
        }
