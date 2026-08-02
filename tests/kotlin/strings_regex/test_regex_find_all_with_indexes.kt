// vybe-test: kotlin/strings_regex/test_regex_find_all_with_indexes
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("[A-Za-z]")
            val first = pattern.find("A1b2C3")
            __check((first?.value ?: "none").toString(), "A")
            val all = pattern.findAll("A1b2C3").toList()
            __check((all[0].range.start).toString(), "0")
            __check((all[1].range.start).toString(), "2")
            __check((all[2].range.start).toString(), "4")
        }
