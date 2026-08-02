// vybe-test: kotlin/property_accessors/test_property_getter_with_range_check
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Counter {
            var value = 0
            val isSmall: Boolean get() = value in 0..10
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter()
            __check((c.isSmall).toString(), "true")
            c.value = 15
            __check((c.isSmall).toString(), "false")
        }
