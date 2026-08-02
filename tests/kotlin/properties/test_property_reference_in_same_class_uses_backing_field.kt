// vybe-test: kotlin/properties/test_property_reference_in_same_class_uses_backing_field
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Counter {
            private var raw = 0

            var value: Int
                get() = raw
                set(next) { raw = next }

            fun bump() {
                value++
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = Counter()
            counter.bump()
            counter.bump()
            __check((counter.value).toString(), "2")
        }
