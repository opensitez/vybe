// vybe-test: kotlin/strings/test_index_of_and_last_index_of
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "banana"
            __check((word.indexOf("na")).toString(), "2")
            __check((word.lastIndexOf("na")).toString(), "4")
            __check((word.indexOf("na", 3)).toString(), "4")
            __check((word.indexOf("x")).toString(), "-1")
        }
