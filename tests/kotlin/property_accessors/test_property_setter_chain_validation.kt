// vybe-test: kotlin/property_accessors/test_property_setter_chain_validation
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var value: Int = 1
                set(v) {
                    field = v
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
            b.value = 4
            b.value = b.value + 1
            __check((b.value).toString(), "5")
        }
