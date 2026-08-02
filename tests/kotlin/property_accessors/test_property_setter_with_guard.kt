// vybe-test: kotlin/property_accessors/test_property_setter_with_guard
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var value: Int = 0
                set(v) {
                    field = if (v == 13) 0 else v
                }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            b.value = 13
            __check((b.value).toString(), "0")
        }
