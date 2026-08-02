// vybe-test: kotlin/strings/test_string_builder_mutability
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val builder = StringBuilder()
            builder.append("a")
            builder.append("-")
            builder.append("z")
            builder.insert(2, "middle")
            __check((builder.toString()).toString(), "a-middlez")
            __check((builder.length).toString(), "9")
        }
