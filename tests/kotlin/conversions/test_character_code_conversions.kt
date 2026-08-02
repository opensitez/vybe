// vybe-test: kotlin/conversions/test_character_code_conversions
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "AZ"
            val first = text[0]
            val second = text[1]
            val asString = first.toString() + second.toString()
            __check((asString).toString(), "AZ")
            __check((first.code).toString(), "65")
            __check((second.code).toString(), "90")
        }
