// vybe-test: kotlin/kotlin_regex_advanced/test_regex_capture_groups_and_names
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val regex = Regex("^(?<name>[a-z]+):(?<value>\\d+)$")
            val result = regex.find("age:42")
            __check((result != null).toString(), "true")
            val groups = result?.groups
            __check((groups?.size).toString(), "3")
            __check((groups?.get("name")?.value).toString(), "age")
            __check((groups?.get("value")?.value).toString(), "42")
        }
