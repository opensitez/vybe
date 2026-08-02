// vybe-test: kotlin/strings/test_trim_start_and_end_variants
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "  Kotlin  "
            __check((value.trimStart()).toString(), "Kotlin  ")
            __check((value.trimEnd()).toString(), "  Kotlin")
            __check((value.trim()).toString(), "Kotlin")
        }
