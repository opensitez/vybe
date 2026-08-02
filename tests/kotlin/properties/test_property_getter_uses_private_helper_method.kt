// vybe-test: kotlin/properties/test_property_getter_uses_private_helper_method
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Formatter {
            private var raw = "kotlin"
            val formatted: String
                get() = raw.uppercase()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val formatter = Formatter()
            __check((formatter.formatted).toString(), "KOTLIN")
        }
