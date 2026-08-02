// vybe-test: kotlin/conversions/test_number_to_string_preserves_roundtrip_for_float_and_double
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 1.0f
            val b = 2.5f
            val d = 3.5
            val fromStringToFloat = a.toString().toFloat()
            val fromStringToDouble = d.toString().toDouble()
            __check((a.toString()).toString(), "1.0")
            __check((fromStringToFloat).toString(), "1.0")
            __check((fromStringToDouble).toString(), "3.5")
            __check((b.toString()).toString(), "2.5")
            __check((d.toString()).toString(), "3.5")
        }
