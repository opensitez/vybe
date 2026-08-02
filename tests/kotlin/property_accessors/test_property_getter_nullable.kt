// vybe-test: kotlin/property_accessors/test_property_getter_nullable
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var text: String? = null
            val safe: String get() = text ?: "none"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().safe).toString(), "none")
            val b = Box()
            b.text = "x"
            __check((b.safe).toString(), "x")
        }
