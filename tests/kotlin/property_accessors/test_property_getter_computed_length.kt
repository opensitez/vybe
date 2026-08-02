// vybe-test: kotlin/property_accessors/test_property_getter_computed_length
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Text {
            var value: String = ""
                set(v) { field = v }
            val length: Int get() = value.length
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Text()
            t.value = "abc"
            __check((t.length).toString(), "3")
        }
