// vybe-test: kotlin/properties/test_top_level_property_read
// origin: languages/kotlin/tests/kotlin/test_properties.rs

val welcome = "hello"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((welcome).toString(), "hello")
        }
