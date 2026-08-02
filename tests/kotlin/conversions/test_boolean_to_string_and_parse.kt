// vybe-test: kotlin/conversions/test_boolean_to_string_and_parse
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val flag = true
            val text = flag.toString()
            val truthy = "true".toBoolean()
            val falsy = "false".toBoolean()
            __check((text).toString(), "true")
            __check((truthy).toString(), "true")
            __check((falsy).toString(), "false")
        }
