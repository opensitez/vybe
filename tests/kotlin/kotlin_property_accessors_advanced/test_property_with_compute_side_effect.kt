// vybe-test: kotlin/kotlin_property_accessors_advanced/test_property_with_compute_side_effect
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Toggle {
            private var on = false
            val flag: Boolean
                get() {
                    on = !on
                    return on
                }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Toggle()
            __check((t.flag).toString(), "true")
            __check((t.flag).toString(), "false")
        }
