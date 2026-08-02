// vybe-test: kotlin/property_accessors/test_property_backed_by_computed_expression
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var a: Int = 2
            var b: Int = 3
            val sum: Int
                get() = a + b
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            __check((b.sum).toString(), "5")
            b.a = 5
            __check((b.sum).toString(), "8")
        }
