// vybe-test: kotlin/kotlin_property_accessors_advanced/test_setter_with_previous_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Running {
            private var _sum = 0
            var sum: Int
                get() = _sum
                set(value) { _sum += value }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = Running()
            r.sum = 3
            r.sum = 4
            __check((r.sum).toString(), "7")
        }
