// vybe-test: kotlin/kotlin_regex_advanced/test_regex_replace_with_transform
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val regex = Regex("(\\w)(\\d+)")
            val replaced = regex.replace("a1 b22", { match ->
                match.groupValues[1] + "=" + match.groupValues[2]
            })
            __check((replaced).toString(), "a=1 b=22")
        }
