// vybe-test: kotlin/property_accessors/test_property_getter_boolean_logic
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Status(val active: Boolean) {
            val activeText: String get() = if (active) "yes" else "no"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Status(true).activeText).toString(), "yes")
            __check((Status(false).activeText).toString(), "no")
        }
