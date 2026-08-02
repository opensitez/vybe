// vybe-test: kotlin/kotlin_property_accessors_advanced/test_backing_field_visible_to_accessors
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Name {
            var text: String = "x"
                set(value) {
                    field = value.uppercase()
                }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = Name()
            n.text = "ab"
            __check((n.text).toString(), "AB")
        }
