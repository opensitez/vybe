// vybe-test: kotlin/property_accessors/test_property_array_accessor
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Bag {
            private val values = IntArray(3)
            var second: Int
                get() = values[1]
                set(v) { values[1] = v }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Bag()
            b.second = 7
            __check((b.second).toString(), "7")
        }
