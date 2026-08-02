// vybe-test: kotlin/property_accessors/test_property_delegates_to_method
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var value: Int = 3
            val visible: Int get() = display()
            private fun display() = value
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().visible).toString(), "3")
        }
