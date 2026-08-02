// vybe-test: kotlin/property_accessors/test_property_private_setter
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var value: Int = 1
                private set
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            __check((b.value).toString(), "1")
        }
