// vybe-test: kotlin/properties/test_top_level_mutable_property_is_shared_across_calls
// origin: languages/kotlin/tests/kotlin/test_properties.rs

var total = 0

        fun inc() {
            total += 1
        }

        fun snapshot(): Int = total

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((snapshot()).toString(), "0")
            inc()
            __check((snapshot()).toString(), "1")
            inc()
            __check((snapshot()).toString(), "2")
        }
