// vybe-test: kotlin/properties/test_property_overridden_accessor_updates_backing_field
// origin: languages/kotlin/tests/kotlin/test_properties.rs

interface ValueSource {
            var value: Int
        }

        class Wrapper : ValueSource {
            private var raw = 4
            override var value: Int
                get() = raw
                set(next) { raw = next - 2 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: ValueSource = Wrapper()
            value.value = 10
            __check((value.value).toString(), "8")
        }
