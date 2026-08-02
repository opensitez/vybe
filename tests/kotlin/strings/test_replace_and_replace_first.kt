// vybe-test: kotlin/strings/test_replace_and_replace_first
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "banana"
            __check((word.replace("na", "NA")).toString(), "baNANA")
            __check((word.replaceFirst("ba", "BO")).toString(), "NAna")
            __check((word.replace("na", "", false)).toString(), "baa")
        }
