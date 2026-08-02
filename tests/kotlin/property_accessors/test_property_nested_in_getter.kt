// vybe-test: kotlin/property_accessors/test_property_nested_in_getter
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var v = 1
            val double: Int get() = run {
                val m = v * 2
                m
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().double).toString(), "2")
        }
