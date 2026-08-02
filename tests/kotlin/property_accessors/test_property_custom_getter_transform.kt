// vybe-test: kotlin/property_accessors/test_property_custom_getter_transform
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box(val raw: String) {
            val trimmed: String get() = raw.trim()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box("  x ").trimmed).toString(), "x")
        }
