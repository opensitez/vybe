// vybe-test: kotlin/kotlin_property_accessors_advanced/test_custom_accessor_computes_transient_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Meter {
            private var ticks = 0
            var total: Int
                get() = ticks + 1
                set(v) { ticks = v }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = Meter()
            m.total = 3
            __check((m.total).toString(), "4")
        }
