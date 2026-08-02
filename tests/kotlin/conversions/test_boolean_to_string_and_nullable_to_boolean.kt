// vybe-test: kotlin/conversions/test_boolean_to_string_and_nullable_to_boolean
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = false.toString()
            val b = true.toString()
            val maybe: Boolean? = null
            __check((a).toString(), "false")
            __check((b).toString(), "true")
            __check((maybe?.toString() ?: "null").toString(), "null")
        }
