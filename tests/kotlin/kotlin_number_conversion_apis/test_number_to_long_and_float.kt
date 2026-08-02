// vybe-test: kotlin/kotlin_number_conversion_apis/test_number_to_long_and_float
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3_000_000_000L.toInt()).toString(), "-1294967296")
            __check((10L.toDouble()).toString(), "10000000000.0")
            __check((42L.toFloat()).toString(), "42.0")
            __check((42L.toByte()).toString(), "42")
        }
