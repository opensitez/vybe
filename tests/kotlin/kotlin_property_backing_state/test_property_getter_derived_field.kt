// vybe-test: kotlin/kotlin_property_backing_state/test_property_getter_derived_field
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_backing_state.rs

class Timer {
            private var total = 0
            var seconds: Int
                get() = total
                set(value) {
                    total = value
                }
            val isZero: Boolean
                get() = total == 0
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Timer()
            __check((t.isZero).toString(), "true")
            t.seconds = 3
            __check((t.seconds).toString(), "3")
            __check((t.isZero).toString(), "false")
        }
