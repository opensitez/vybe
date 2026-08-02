// vybe-test: kotlin/properties/test_late_mutation_after_top_level_property_read
// origin: languages/kotlin/tests/kotlin/test_properties.rs

var current = 1

        fun bump() {
            current += 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((current).toString(), "1")
            bump()
            __check((current).toString(), "3")
            current = 10
            __check((current).toString(), "10")
        }
