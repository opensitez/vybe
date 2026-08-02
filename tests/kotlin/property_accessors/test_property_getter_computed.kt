// vybe-test: kotlin/property_accessors/test_property_getter_computed
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box(val a: Int, val b: Int) {
            val sum: Int get() = a + b
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box(2, 3).sum).toString(), "5")
        }
