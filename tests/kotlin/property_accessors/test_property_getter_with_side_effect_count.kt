// vybe-test: kotlin/property_accessors/test_property_getter_with_side_effect_count
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var c = 0
            val label: Int
                get() {
                    c += 1
                    return c
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
            __check((b.label).toString(), "1")
            __check((b.label).toString(), "2")
        }
