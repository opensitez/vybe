// vybe-test: kotlin/properties/test_property_setter_updates_derived_backing_value
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Range {
            private var current: Int = 0
            var base: Int
                get() = current
                set(next) { current = if (next > 100) 100 else next }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = Range()
            r.base = 150
            __check((r.base).toString(), "100")
        }
