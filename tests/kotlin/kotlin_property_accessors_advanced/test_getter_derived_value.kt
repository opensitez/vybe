// vybe-test: kotlin/kotlin_property_accessors_advanced/test_getter_derived_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Holder {
            private val value = 4
            val doubled get() = value * 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder().doubled).toString(), "8")
        }
