// vybe-test: kotlin/property_accessors/test_property_delayed_init
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            val v: Int by lazy { 5 }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            __check((b.v).toString(), "5")
            __check((b.v).toString(), "5")
        }
