// vybe-test: kotlin/basic/test_string_template
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name = "Kotlin"
            val version = 1
            __check(("Language ${name} ${version + 1}").toString(), "Language Kotlin 2")
        }
