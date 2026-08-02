// vybe-test: kotlin/property_accessors/test_property_getter_calls_method
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            val a = 1
            val b = 2
            val total: Int get() = sum()
            fun sum() = a + b
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().total).toString(), "3")
        }
