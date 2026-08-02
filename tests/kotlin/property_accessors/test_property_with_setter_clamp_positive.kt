// vybe-test: kotlin/property_accessors/test_property_with_setter_clamp_positive
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Clamp {
            var value: Int = 0
                set(v) {
                    field = if (v < 0) 0 else v
                }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Clamp()
            c.value = -1
            __check((c.value).toString(), "0")
        }
