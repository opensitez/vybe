// vybe-test: kotlin/strings/test_string_replace_range_and_region
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "kotlin"
            __check((value.replaceRange(1, 4, "A")).toString(), "kAin")
            __check((value.replaceRange(0, 1, "Z")).toString(), "Zotlin")
            __check((value.replaceFirst("li", "LI")).toString(), "kotLIn")
        }
