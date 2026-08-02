// vybe-test: kotlin/strings/test_case_transformations
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "Kotlin"
            __check((value.lowercase()).toString(), "kotlin")
            __check((value.uppercase()).toString(), "KOTLIN")
        }
