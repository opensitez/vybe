// vybe-test: kotlin/strings_regex/test_regex_named_capture_group_access
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("(?<id>\\d+)-(?<name>\\w+)")
            val result = pattern.find("42-kotlin")
            __check((result?.groups?.get("id")?.value ?: "missing").toString(), "42")
            __check((result?.groups?.get("name")?.value ?: "missing").toString(), "kotlin")
        }
