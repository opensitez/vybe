// vybe-test: kotlin/printing/test_printing_printing_of_boolean_array_contents
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val flags = booleanArrayOf(true, false, true)
            __check((flags.joinToString(",")).toString(), "true,false,true")
        }
