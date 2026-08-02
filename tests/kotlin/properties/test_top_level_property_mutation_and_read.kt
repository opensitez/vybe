// vybe-test: kotlin/properties/test_top_level_property_mutation_and_read
// origin: languages/kotlin/tests/kotlin/test_properties.rs

var score = 0

        fun inc() {
            score += 5
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score).toString(), "0")
            inc()
            __check((score).toString(), "5")
        }
